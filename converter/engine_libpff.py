"""libpff (pypff) tabanli OST okuyucu / disa aktarici.

Bu motor, diskteki bir ``.ost`` dosyasini DOGRUDAN okur; dosyanin bir Outlook
profiline bagli olmasi gerekmez (orphan / eski yedek OST dosyalari dahil).

Yetenekler:
    * ``build_tree``          : OST'nin klasor agacini (govde okumadan, hizli)
      cikarir. Arayuz bu agaci kullaniciya gosterir; kullanici hangi klasor /
      mesajlarin PST'ye gececegini secer.
    * ``list_folder_messages``: Tek bir klasorun mesaj basliklarini (konu,
      gonderen, tarih) tembel (lazy) olarak listeler.
    * ``export_selected_eml`` : SECILEN klasor/mesajlari, OST'deki klasor
      yapisini BIREBIR koruyarak ``.eml`` dosyalari halinde yazar. Her klasor
      icin gercek adi bir yan dosyada (``__foldername__.txt``) saklar; boylece
      Outlook'a aktarimda klasor adlari ve hiyerarsi aynen olusur.
    * ``export_selected_mbox``: Secilen mesajlari tek bir MBOX dosyasina yazar.

Onceki klasor-bozulmasi sorunlarinin cozumu:
    * Kok sarmalayici klasorler ("Top of Personal Folders" / IPM_SUBTREE) artik
      DUZLESTIRILIR; gercek kullanici klasorleri (Gelen Kutusu, Gonderilmis...)
      PST kokunde, OST'deki gibi gorunur.
    * Kardes klasor adlari sanitize sonrasi cakisirsa benzersizlestirilir; klasor
      adlari kirpilmadan, yan dosya araciligi ile birebir korunur.
    * Her mesaja ``Date`` (ve varsa orijinal basliklar) eklenir; Outlook'ta
      mailler tarihli ve okunabilir gorunur.

Gereksinim: ``pip install libpff-python``  (Windows'ta onceden derlenmis bir
tekerlek/wheel gerekebilir). Kurulu degilse ``is_available()`` False doner.
"""

from __future__ import annotations

import datetime as _dt
import os
import re
from email.message import EmailMessage
from email.parser import Parser
from email.policy import default as default_policy
from email.utils import format_datetime
from typing import Callable, Dict, List, Optional, Set, Tuple, Union

#: Her disa aktarilan klasor diziniine yazilan, klasorun GERCEK adini tutan
#: yan dosya. Outlook'a aktarimda (engine_outlook) bu ad birebir kullanilir.
FOLDERNAME_FILE = "__foldername__.txt"

#: Kok sarmalayici klasorler - bunlar PST kokune yansitilMAZ; cocuklari bir ust
#: seviyeye tasinir. Boylece OST'deki gercek yapi (Gelen Kutusu vb.) korunur.
_ROOT_CONTAINERS = {
    "",
    "adsiz",
    "top of personal folders",
    "top of outlook data file",
    "ipm_subtree",
    "root",
    "root - mailbox",
    "root - public folders",
    "kok",
    "kok - posta kutusu",
    "ust klasor",
    "ust bilgi",
}

#: Bir folder_id; kok altindaki ham libpff alt-klasor indekslerinin "/" ile
#: birlestirilmis halidir (orn. "0/2/1"). build_tree ve disa aktarim ayni
#: gezinmeyi yaptigi icin bu kimlikler iki tarafta da ESLESIR.
Selection = Dict[str, Union[bool, Set[int]]]


def is_available() -> bool:
    try:
        import pypff  # noqa: F401
    except Exception:
        return False
    return True


# --------------------------------------------------------------------------- #
# Hafif veri modelleri
# --------------------------------------------------------------------------- #
class Attachment:
    def __init__(self, name: str, data: bytes):
        self.name = name
        self.data = data


class Message:
    def __init__(self):
        self.subject: str = ""
        self.sender_name: str = ""
        self.headers: str = ""
        self.plain_body: str = ""
        self.html_body: str = ""
        self.date = None  # datetime veya None
        self.attachments: List[Attachment] = []


class Folder:
    def __init__(self, name: str):
        self.name = name
        self.subfolders: List["Folder"] = []
        self.messages: List[Message] = []

    def total_messages(self) -> int:
        return len(self.messages) + sum(f.total_messages() for f in self.subfolders)


# --------------------------------------------------------------------------- #
# Okuma
# --------------------------------------------------------------------------- #
def _safe(getter, default=""):
    try:
        val = getter()
        return val if val is not None else default
    except Exception:
        return default


def _message_date(pff_msg):
    """Mesajin gonderim/teslim zamanini (datetime) en iyi cabayla dondurur."""
    for attr in ("client_submit_time", "delivery_time",
                 "creation_time", "modification_time"):
        val = _safe(lambda a=attr: getattr(pff_msg, a), None)
        if isinstance(val, _dt.datetime):
            return val
    return None


def _read_message(pff_msg) -> Message:
    msg = Message()

    msg.subject = _safe(lambda: pff_msg.subject) or "(konusuz)"
    msg.sender_name = _safe(lambda: pff_msg.sender_name)
    msg.headers = _safe(lambda: pff_msg.transport_headers)
    msg.plain_body = _safe(lambda: pff_msg.plain_text_body)
    msg.html_body = _safe(lambda: pff_msg.html_body)
    msg.date = _message_date(pff_msg)

    # Bazi pypff surumleri bytes dondurur.
    for attr in ("plain_body", "html_body", "headers"):
        val = getattr(msg, attr)
        if isinstance(val, bytes):
            setattr(msg, attr, val.decode("utf-8", "replace"))

    try:
        for i in range(pff_msg.number_of_attachments):
            att = pff_msg.get_attachment(i)
            data = b""
            try:
                size = att.get_size()
                data = att.read_buffer(size)
            except Exception:
                try:
                    data = att.read_buffer(att.size)
                except Exception:
                    data = b""
            name = "ek.bin"
            for getter in ("get_name", "name"):
                try:
                    candidate = getattr(att, getter)
                    candidate = candidate() if callable(candidate) else candidate
                    if candidate:
                        name = candidate
                        break
                except Exception:
                    continue
            msg.attachments.append(Attachment(str(name), data or b""))
    except Exception:
        pass

    return msg


def _open_pff(path: str):
    """OST/PST dosyasini acar ve (pff, file_obj) dondurur.

    Windows'ta iki ayri sorunu birden asar:

    * **Ileri egik cizgi**: tkinter yollari ``C:/Users/...`` biciminde verir.
      libpff, ``\\\\?\\`` uzun-yol bicimini yalnizca TERS egik cizgi ile kabul
      eder; ileri cizgi gorunce surucu harfini kaybedip yolu bozar. Yolu once
      ``os.path.normpath`` ile yerel ayraca (``\\``) ceviririz.
    * **Unicode / Turkce karakter**: Dosyayi once Python'un actigi bir dosya
      tutamaci (``open_file_object``) ile vermeyi deneriz; Unicode-guvenlidir.

    Iki yontem de denenir; her ikisi de basarisiz olursa ayrintili hata verilir.
    Cagiran taraf isi bitince ``pff.close()`` ve (varsa) ``file_obj.close()``
    cagirmalidir.
    """
    if not is_available():
        raise RuntimeError("libpff (pypff) bulunamadi. 'pip install libpff-python'")
    import pypff

    if not os.path.exists(path):
        raise FileNotFoundError(path)

    native = os.path.normpath(os.fspath(path))
    errors = []

    # 1) Python dosya tutamaci ile (Unicode & ayrac guvenli) - tercih edilen.
    file_obj = None
    pff = pypff.file()
    try:
        file_obj = open(native, "rb")
        pff.open_file_object(file_obj)
        return pff, file_obj
    except Exception as exc:
        errors.append("dosya-tutamaci: %s" % exc)
        try:
            pff.close()
        except Exception:
            pass
        if file_obj is not None:
            try:
                file_obj.close()
            except Exception:
                pass

    # 2) Yerel (ters egik cizgili) yol ile dogrudan ac.
    pff = pypff.file()
    try:
        pff.open(native)
        return pff, None
    except Exception as exc:
        errors.append("yol: %s" % exc)
        try:
            pff.close()
        except Exception:
            pass

    raise RuntimeError(
        "OST dosyasi acilamadi.\nYol: %s\nDenemeler:\n  - %s"
        % (native, "\n  - ".join(errors))
    )


def _close_pff(pff, file_obj) -> None:
    try:
        if pff is not None:
            pff.close()
    except Exception:
        pass
    if file_obj is not None:
        try:
            file_obj.close()
        except Exception:
            pass


def _folder_name(pff_folder, default: str = "Adsiz") -> str:
    try:
        return pff_folder.name or default
    except Exception:
        return default


def _is_container(name: str) -> bool:
    return (name or "").strip().lower() in _ROOT_CONTAINERS


# --------------------------------------------------------------------------- #
# Klasor agaci (govde okumadan, hizli) - arayuzdeki secim agaci icin
# --------------------------------------------------------------------------- #
def _node_list(root) -> List[dict]:
    """Kok klasorden baslayarak gorunur klasor dugumlerini DFS on-sira dondurur.

    Her dugum: ``{"id", "parent", "name", "count"}``.
    ``id``  : ham libpff indeks yolu ("0/2/1").
    ``parent``: ust GORUNUR dugumun id'si (kok sarmalayicilar duzlestirildigi
              icin gercek ebeveynden farkli olabilir) veya None (kok seviye).
    ``count``: klasordeki dogrudan mesaj sayisi.

    Kok sarmalayici klasorler (``_is_container``) bir dugum URETMEZ; cocuklari
    bir ust gorunur seviyeye baglanir.
    """
    nodes: List[dict] = []

    def rec(folder, path: Tuple[int, ...], parent_id: Optional[str]) -> None:
        try:
            n_sub = folder.number_of_sub_folders
        except Exception:
            n_sub = 0
        for i in range(n_sub):
            try:
                sub = folder.get_sub_folder(i)
            except Exception:
                continue
            spath = path + (i,)
            sid = "/".join(str(p) for p in spath)
            name = _folder_name(sub)
            if _is_container(name):
                # Sarmalayici: dugum olusturma, cocuklari ust seviyeye tasi.
                rec(sub, spath, parent_id)
                continue
            try:
                count = sub.number_of_sub_messages
            except Exception:
                count = 0
            nodes.append({"id": sid, "parent": parent_id, "name": name,
                          "count": int(count)})
            rec(sub, spath, sid)

    rec(root, (), None)
    return nodes


def build_tree(path: str) -> List[dict]:
    """OST'nin klasor agacini (govde okumadan) hizlica cikarir.

    Donus: ``_node_list`` bicimi dugum listesi (DFS on-sira).
    """
    pff, file_obj = _open_pff(path)
    try:
        return _node_list(pff.get_root_folder())
    finally:
        _close_pff(pff, file_obj)


def _folder_by_id(root, fid: str):
    """"0/2/1" bicimli kimlikten ilgili libpff klasorune iner."""
    folder = root
    if fid:
        for part in fid.split("/"):
            folder = folder.get_sub_folder(int(part))
    return folder


def _fmt_date_short(d) -> str:
    if isinstance(d, _dt.datetime):
        try:
            return d.strftime("%Y-%m-%d %H:%M")
        except Exception:
            return ""
    return str(d) if d else ""


def list_folder_messages(path: str, fid: str) -> List[dict]:
    """Tek bir klasorun mesaj basliklarini (konu/gonderen/tarih) listeler.

    Govdeler okunmaz; arayuzde klasor genisletildiginde tembel cagrilir.
    Donus ogeleri: ``{"index", "subject", "sender", "date"}`` (index 0-tabanli).
    """
    pff, file_obj = _open_pff(path)
    try:
        folder = _folder_by_id(pff.get_root_folder(), fid)
        out: List[dict] = []
        try:
            n = folder.number_of_sub_messages
        except Exception:
            n = 0
        for i in range(n):
            try:
                m = folder.get_sub_message(i)
            except Exception:
                out.append({"index": i, "subject": "(okunamadi)",
                            "sender": "", "date": ""})
                continue
            out.append({
                "index": i,
                "subject": (_safe(lambda: m.subject) or "(konusuz)"),
                "sender": _safe(lambda: m.sender_name),
                "date": _fmt_date_short(_message_date(m)),
            })
        return out
    finally:
        _close_pff(pff, file_obj)


# --------------------------------------------------------------------------- #
# EML uretimi
# --------------------------------------------------------------------------- #
_INVALID = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def _safe_name(name: str, fallback: str = "oge") -> str:
    name = _INVALID.sub("_", (name or "").strip()) or fallback
    return name[:120]


#: Govde/MIME ile ilgili basliklar - bunlari KENDIMIZ uretecegimiz icin
#: orijinal transport_headers'tan kopyalamayiz (aksi halde "set_content not
#: valid on multipart" hatasi olusur).
_SKIP_HEADERS = {
    "content-type",
    "content-transfer-encoding",
    "content-disposition",
    "content-id",
    "content-description",
    "content-length",
    "mime-version",
}


def message_to_eml_bytes(msg: Message) -> bytes:
    """Bir Message nesnesini RFC822 (.eml) baytlarina cevirir."""
    try:
        return _build_eml(msg)
    except Exception:
        # Hicbir mesaj kaybolmasin: en sade bicimde yeniden kur.
        try:
            fallback = EmailMessage()
            fallback["Subject"] = msg.subject or "(konusuz)"
            if msg.sender_name:
                fallback["From"] = msg.sender_name
            body = msg.plain_body or msg.html_body or " "
            fallback.set_content(body)
            return fallback.as_bytes()
        except Exception:
            # Son care: ham metin.
            raw = "Subject: %s\r\n\r\n%s" % (
                msg.subject or "(konusuz)", msg.plain_body or ""
            )
            return raw.encode("utf-8", "replace")


def _build_eml(msg: Message) -> bytes:
    eml = EmailMessage()

    if msg.headers.strip():
        # Orijinal aciklayici basliklari (From/To/Subject/Date...) koru;
        # govde/MIME basliklarini atla (govdeyi biz kuracagiz).
        parsed = Parser(policy=default_policy).parsestr(msg.headers, headersonly=True)
        for key, value in parsed.items():
            if key.lower() in _SKIP_HEADERS:
                continue
            try:
                eml[key] = value
            except Exception:
                continue

    if "Subject" not in eml:
        eml["Subject"] = msg.subject or "(konusuz)"
    if "From" not in eml and msg.sender_name:
        try:
            eml["From"] = msg.sender_name
        except Exception:
            pass
    if "Date" not in eml and msg.date is not None:
        try:
            d = msg.date
            eml["Date"] = format_datetime(d) if isinstance(d, _dt.datetime) else str(d)
        except Exception:
            pass

    if msg.html_body:
        eml.set_content(msg.plain_body or " ")
        eml.add_alternative(msg.html_body, subtype="html")
    else:
        eml.set_content(msg.plain_body or " ")

    for att in msg.attachments:
        try:
            eml.add_attachment(
                att.data,
                maintype="application",
                subtype="octet-stream",
                filename=_safe_name(att.name, "ek.bin"),
            )
        except Exception:
            continue

    return eml.as_bytes()


class _Progress:
    """Ilerleme/hata raporlamayi kisar: arayuzu binlerce satirla bogmaz."""

    def __init__(self, total, callback, verb="aktarildi", max_fail_logs=15):
        self.total = max(1, total)
        self.cb = callback
        self.verb = verb
        self.done = 0
        self.fail = 0
        self._fail_logged = 0
        self._max_fail_logs = max_fail_logs
        # En fazla ~200 ilerleme guncellemesi -> akici ama tasmayan arayuz.
        self.step = max(50, self.total // 200)

    def _emit(self, msg, frac=-1.0):
        if self.cb:
            self.cb(msg, frac)

    def tick(self):
        self.done += 1
        if self.done % self.step == 0 or self.done >= self.total:
            self._emit("%d/%d mesaj %s" % (self.done, self.total, self.verb),
                       min(1.0, self.done / self.total))

    def failure(self, label, exc):
        self.fail += 1
        if self._fail_logged < self._max_fail_logs:
            self._fail_logged += 1
            self._emit("  ! Atlandi: %s (%s)" % (label, exc))
        elif self._fail_logged == self._max_fail_logs:
            self._fail_logged += 1
            self._emit("  ! (daha fazla atlanan oge sessizce gecilecek)")

    def finish(self):
        if self.fail:
            self._emit("Tamamlandi. %d/%d oge yazildi, %d oge atlandi."
                       % (self.done - self.fail, self.done, self.fail), 1.0)
        else:
            self._emit("Bitti. %d mesaj." % self.done, 1.0)


# --------------------------------------------------------------------------- #
# Secime dayali disa aktarim
# --------------------------------------------------------------------------- #
def _full_selection(nodes: List[dict]) -> Selection:
    """Tum klasorleri (tum mesajlariyla) iceren secim uretir."""
    return {n["id"]: True for n in nodes}


def _resolve_dirs(nodes: List[dict], selection: Selection):
    """Secilen dugumler ve atalari icin benzersiz dizin yollarini hesaplar.

    Donus: ``(byid, needed, relpath)``
      byid    : id -> dugum
      needed  : olusturulmasi gereken dizin id kumesi (secilenler + atalari)
      relpath : id -> ``out_dir`` icindeki goreli dizin yolu
    """
    byid = {n["id"]: n for n in nodes}

    # Secilen her klasor ve TUM atalari materyallesir (ic ice yapi icin).
    needed: Set[str] = set()
    for fid in selection:
        cur: Optional[str] = fid
        while cur is not None and cur in byid and cur not in needed:
            needed.add(cur)
            cur = byid[cur]["parent"]

    # Kardesler arasi benzersiz dizin adi (cakismalari ayikla).
    used: Dict[Optional[str], Set[str]] = {}
    dirname: Dict[str, str] = {}
    for n in nodes:  # DFS on-sira: ebeveyn daima cocuktan once islenir.
        nid = n["id"]
        if nid not in needed:
            continue
        base = _safe_name(n["name"], "klasor")[:80]
        siblings = used.setdefault(n["parent"], set())
        name = base
        k = 2
        while name.lower() in siblings:
            name = "%s_%d" % (base, k)
            k += 1
        siblings.add(name.lower())
        dirname[nid] = name

    relpath: Dict[str, str] = {}
    for n in nodes:
        nid = n["id"]
        if nid not in needed:
            continue
        parent = n["parent"]
        prefix = relpath.get(parent, "")
        relpath[nid] = os.path.join(prefix, dirname[nid]) if prefix else dirname[nid]

    return byid, needed, relpath


def _selected_indices(folder, sel: Union[bool, Set[int]]) -> List[int]:
    try:
        n_msg = folder.number_of_sub_messages
    except Exception:
        n_msg = 0
    if sel is True:
        return list(range(n_msg))
    return sorted(i for i in sel if 0 <= i < n_msg)


def export_selected_eml(
    path: str,
    out_dir: str,
    selection: Optional[Selection] = None,
    progress: Optional[Callable[[str, float], None]] = None,
    short_names: bool = False,
) -> str:
    """Secilen klasor/mesajlari, OST yapisini koruyarak .eml olarak yazar.

    ``selection`` None ise tum posta kutusu disa aktarilir. Her olusturulan
    klasor dizinine, klasorun gercek adini tutan ``__foldername__.txt`` yazilir;
    boylece Outlook'a aktarimda klasor adlari ve hiyerarsi birebir olusur.

    ``short_names=True`` ise dosyalar 00001.eml gibi kisa adlarla yazilir
    (PST'ye aktarim ardisinda Windows MAX_PATH sorununu onler).
    """
    if progress:
        progress("OST aciliyor...", 0.0)
    pff, file_obj = _open_pff(path)
    try:
        root = pff.get_root_folder()
        nodes = _node_list(root)
        if selection is None:
            selection = _full_selection(nodes)

        byid, needed, relpath = _resolve_dirs(nodes, selection)

        total = 0
        for fid, sel in selection.items():
            if fid not in byid:
                continue
            total += byid[fid]["count"] if sel is True else len(sel)
        prog = _Progress(total, progress, verb="aktarildi")

        os.makedirs(out_dir, exist_ok=True)
        # Gerekli tum dizinleri (atalar dahil) yan-dosyalariyla olustur.
        for nid in needed:
            d = os.path.join(out_dir, relpath[nid])
            try:
                os.makedirs(d, exist_ok=True)
                with open(os.path.join(d, FOLDERNAME_FILE), "w",
                          encoding="utf-8") as fh:
                    fh.write(byid[nid]["name"])
            except Exception:
                pass

        for fid, sel in selection.items():
            if fid not in byid:
                continue
            try:
                folder = _folder_by_id(root, fid)
            except Exception as exc:
                prog._emit("  ! Klasor acilamadi (%s): %s" % (fid, exc))
                continue
            dest = os.path.join(out_dir, relpath[fid])
            for i in _selected_indices(folder, sel):
                label = "oge %d" % (i + 1)
                try:
                    msg = _read_message(folder.get_sub_message(i))
                    if short_names:
                        fname = "%05d.eml" % (i + 1)
                    else:
                        label = msg.subject
                        fname = _safe_name("%05d_%s" % (i + 1, msg.subject),
                                           "%05d_mesaj" % (i + 1)) + ".eml"
                    with open(os.path.join(dest, fname), "wb") as fh:
                        fh.write(message_to_eml_bytes(msg))
                except Exception as exc:
                    prog.failure(label, exc)
                prog.tick()

        prog.finish()
        return out_dir
    finally:
        _close_pff(pff, file_obj)


def export_selected_mbox(
    path: str,
    mbox_path: str,
    selection: Optional[Selection] = None,
    progress: Optional[Callable[[str, float], None]] = None,
) -> str:
    """Secilen mesajlari tek bir MBOX dosyasina yazar (yapi duzlestirilir)."""
    import mailbox

    if progress:
        progress("OST aciliyor...", 0.0)
    pff, file_obj = _open_pff(path)
    mbox = mailbox.mbox(mbox_path)
    try:
        mbox.lock()
        root = pff.get_root_folder()
        nodes = _node_list(root)
        if selection is None:
            selection = _full_selection(nodes)
        byid = {n["id"]: n for n in nodes}

        total = 0
        for fid, sel in selection.items():
            if fid not in byid:
                continue
            total += byid[fid]["count"] if sel is True else len(sel)
        prog = _Progress(total, progress, verb="yazildi")

        for fid, sel in selection.items():
            if fid not in byid:
                continue
            try:
                folder = _folder_by_id(root, fid)
            except Exception:
                continue
            for i in _selected_indices(folder, sel):
                try:
                    msg = _read_message(folder.get_sub_message(i))
                    mbox.add(message_to_eml_bytes(msg))
                except Exception as exc:
                    prog.failure("oge %d" % (i + 1), exc)
                prog.tick()

        mbox.flush()
        prog.finish()
        return mbox_path
    finally:
        try:
            mbox.unlock()
        except Exception:
            pass
        try:
            mbox.close()
        except Exception:
            pass
        _close_pff(pff, file_obj)


# --------------------------------------------------------------------------- #
# Geriye donuk uyumlu sarmalayicilar (tum posta kutusu)
# --------------------------------------------------------------------------- #
def export_eml_tree(
    path: str,
    out_dir: str,
    progress: Optional[Callable[[str, float], None]] = None,
    short_names: bool = False,
) -> str:
    """Tum OST'yi .eml agaci olarak disa aktarir (secim = hepsi)."""
    return export_selected_eml(path, out_dir, None, progress, short_names)


def export_mbox(
    path: str,
    mbox_path: str,
    progress: Optional[Callable[[str, float], None]] = None,
) -> str:
    """Tum OST'yi tek MBOX dosyasina yazar (secim = hepsi)."""
    return export_selected_mbox(path, mbox_path, None, progress)


def read_ost(path: str) -> Folder:
    """OST/PST dosyasini okuyup klasor agacini (bellege) dondurur.

    Not: Buyuk posta kutularinda bellek tuketir; toplu disa aktarim icin
    bunun yerine ``export_selected_eml`` / ``export_selected_mbox`` akis
    fonksiyonlarini kullanin.
    """
    pff, file_obj = _open_pff(path)
    try:
        tree = _read_folder(pff.get_root_folder())
        tree.name = os.path.splitext(os.path.basename(path))[0]
        return tree
    finally:
        _close_pff(pff, file_obj)


def _read_folder(pff_folder) -> Folder:
    folder = Folder(_folder_name(pff_folder))
    try:
        for i in range(pff_folder.number_of_sub_messages):
            folder.messages.append(_read_message(pff_folder.get_sub_message(i)))
    except Exception:
        pass
    try:
        for i in range(pff_folder.number_of_sub_folders):
            folder.subfolders.append(_read_folder(pff_folder.get_sub_folder(i)))
    except Exception:
        pass
    return folder
