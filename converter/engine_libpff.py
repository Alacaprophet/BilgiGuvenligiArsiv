"""libpff (pypff) tabanli OST okuyucu / disa aktarici.

Bu motor, diskteki bir ``.ost`` dosyasini DOGRUDAN okur; dosyanin bir Outlook
profiline bagli olmasi gerekmez (orphan / eski yedek OST dosyalari dahil).

Cikti secenekleri:
    * ``export_eml_tree`` : Her mesaji ``.eml`` olarak, OST'deki klasor yapisini
      koruyarak bir dizine yazar. Bu ciktilar Outlook'a engine_outlook araciligi
      ile aktarilarak GERCEK bir PST'ye donusturulebilir; ayrica tek tek de
      acilabilir.
    * ``export_mbox`` : Tum mesajlari tek bir MBOX dosyasina yazar (Thunderbird
      vb. icin uygundur).

Gereksinim: ``pip install libpff-python``  (Windows'ta onceden derlenmis bir
tekerlek/wheel gerekebilir). Kurulu degilse ``is_available()`` False doner.
"""

from __future__ import annotations

import os
import re
from email.message import EmailMessage
from email.parser import Parser
from email.policy import default as default_policy
from typing import Callable, List, Optional


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
def _read_message(pff_msg) -> Message:
    msg = Message()

    def safe(getter, default=""):
        try:
            val = getter()
            return val if val is not None else default
        except Exception:
            return default

    msg.subject = safe(lambda: pff_msg.subject) or "(konusuz)"
    msg.sender_name = safe(lambda: pff_msg.sender_name)
    msg.headers = safe(lambda: pff_msg.transport_headers)
    msg.plain_body = safe(lambda: pff_msg.plain_text_body)
    msg.html_body = safe(lambda: pff_msg.html_body)

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


def _read_folder(pff_folder) -> Folder:
    name = "Adsiz"
    try:
        name = pff_folder.name or "Adsiz"
    except Exception:
        pass
    folder = Folder(name)

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


def _count_messages(pff_folder) -> int:
    """Govde okumadan, hizlica toplam mesaj sayisini hesaplar (ilerleme icin)."""
    total = 0
    try:
        total += pff_folder.number_of_sub_messages
    except Exception:
        pass
    try:
        for i in range(pff_folder.number_of_sub_folders):
            total += _count_messages(pff_folder.get_sub_folder(i))
    except Exception:
        pass
    return total


def read_ost(path: str) -> Folder:
    """OST/PST dosyasini okuyup klasor agacini (bellege) dondurur.

    Not: Buyuk posta kutularinda bellek tuketir; toplu disa aktarim icin
    bunun yerine ``export_eml_tree`` / ``export_mbox`` akis (streaming)
    fonksiyonlarini kullanin.
    """
    pff, file_obj = _open_pff(path)
    try:
        tree = _read_folder(pff.get_root_folder())
        tree.name = os.path.splitext(os.path.basename(path))[0]
        return tree
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


def export_eml_tree(
    path: str,
    out_dir: str,
    progress: Optional[Callable[[str, float], None]] = None,
    short_names: bool = False,
) -> str:
    """OST'yi klasor yapisini koruyarak .eml dosyalari halinde disa aktarir.

    Akis (streaming) yontemiyle calisir: tum posta kutusunu bellege almak
    yerine her mesaji okuyup hemen yazar -> dusuk bellek, yuksek hiz.

    short_names=True ise dosyalar 00001.eml gibi kisa adlarla yazilir
    (PST'ye aktarim ardisinda Windows MAX_PATH sorununu ve gereksiz isi onler).
    """
    if progress:
        progress("OST aciliyor...", 0.0)
    pff, file_obj = _open_pff(path)
    try:
        root = pff.get_root_folder()
        prog = _Progress(_count_messages(root), progress, verb="aktarildi")
        os.makedirs(out_dir, exist_ok=True)

        def walk(folder, rel: str) -> None:
            dest = os.path.join(out_dir, rel) if rel else out_dir
            try:
                os.makedirs(dest, exist_ok=True)
            except Exception:
                pass
            try:
                n_msg = folder.number_of_sub_messages
            except Exception:
                n_msg = 0
            for i in range(n_msg):
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
            try:
                n_sub = folder.number_of_sub_folders
            except Exception:
                n_sub = 0
            for j in range(n_sub):
                try:
                    sub = folder.get_sub_folder(j)
                except Exception:
                    continue
                sname = _safe_name(_folder_name(sub), "klasor")[:60]
                walk(sub, os.path.join(rel, sname) if rel else sname)

        walk(root, "")
        prog.finish()
        return out_dir
    finally:
        _close_pff(pff, file_obj)


def export_mbox(
    path: str,
    mbox_path: str,
    progress: Optional[Callable[[str, float], None]] = None,
) -> str:
    """Tum mesajlari tek bir MBOX dosyasina yazar (akis yontemiyle)."""
    import mailbox

    if progress:
        progress("OST aciliyor...", 0.0)
    pff, file_obj = _open_pff(path)
    mbox = mailbox.mbox(mbox_path)
    try:
        mbox.lock()
        root = pff.get_root_folder()
        prog = _Progress(_count_messages(root), progress, verb="yazildi")

        def walk(folder) -> None:
            try:
                n_msg = folder.number_of_sub_messages
            except Exception:
                n_msg = 0
            for i in range(n_msg):
                try:
                    msg = _read_message(folder.get_sub_message(i))
                    mbox.add(message_to_eml_bytes(msg))
                except Exception as exc:
                    prog.failure("oge %d" % (i + 1), exc)
                prog.tick()
            try:
                n_sub = folder.number_of_sub_folders
            except Exception:
                n_sub = 0
            for j in range(n_sub):
                try:
                    walk(folder.get_sub_folder(j))
                except Exception:
                    continue

        walk(root)
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
