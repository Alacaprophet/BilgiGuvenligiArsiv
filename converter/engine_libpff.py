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


def read_ost(path: str) -> Folder:
    """OST/PST dosyasini okuyup klasor agacini dondurur.

    Windows'ta iki ayri sorunu birden asar:

    * **Ileri egik cizgi**: tkinter yollari ``C:/Users/...`` biciminde verir.
      libpff, ``\\\\?\\`` uzun-yol bicimini yalnizca TERS egik cizgi ile kabul
      eder; ileri cizgi gorunce surucu harfini kaybedip yolu bozar. Yolu once
      ``os.path.normpath`` ile yerel ayraca (``\\``) ceviririz.
    * **Unicode / Turkce karakter**: Dosyayi once Python'un actigi bir dosya
      tutamaci (``open_file_object``) ile vermeyi deneriz; bu yol Unicode-guvenli
      oldugu icin "Outlook Dosyalari" gibi yollar sorunsuz acilir.

    Iki yontem de denenir; her ikisi de basarisiz olursa ayrintili hata verilir.
    """
    if not is_available():
        raise RuntimeError("libpff (pypff) bulunamadi. 'pip install libpff-python'")
    import pypff

    if not os.path.exists(path):
        raise FileNotFoundError(path)

    name = os.path.splitext(os.path.basename(path))[0]
    # Ileri egik cizgileri yerel ayraca cevir (Windows'ta C:/.. -> C:\..).
    native = os.path.normpath(os.fspath(path))

    errors = []

    # 1) Python dosya tutamaci ile (Unicode & ayrac guvenli) - tercih edilen.
    file_obj = None
    pff = pypff.file()
    try:
        file_obj = open(native, "rb")
        pff.open_file_object(file_obj)
        tree = _read_folder(pff.get_root_folder())
        tree.name = name
        return tree
    except Exception as exc:
        errors.append("dosya-tutamaci: %s" % exc)
    finally:
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
        tree = _read_folder(pff.get_root_folder())
        tree.name = name
        return tree
    except Exception as exc:
        errors.append("yol: %s" % exc)
    finally:
        try:
            pff.close()
        except Exception:
            pass

    raise RuntimeError(
        "OST dosyasi acilamadi.\nYol: %s\nDenemeler:\n  - %s"
        % (native, "\n  - ".join(errors))
    )


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


def export_eml_tree(
    path: str,
    out_dir: str,
    progress: Optional[Callable[[str, float], None]] = None,
) -> str:
    """OST'yi klasor yapisini koruyarak .eml dosyalari halinde disa aktarir.

    Returns: cikti kok dizini.
    """
    tree = read_ost(path)
    os.makedirs(out_dir, exist_ok=True)

    total = max(1, tree.total_messages())
    done = 0

    def report(m: str, f: float = -1.0) -> None:
        if progress:
            progress(m, f)

    def walk(folder: Folder, rel: str) -> None:
        nonlocal done
        dest = os.path.join(out_dir, rel) if rel else out_dir
        os.makedirs(dest, exist_ok=True)
        for i, msg in enumerate(folder.messages, start=1):
            fname = _safe_name(f"{i:05d}_{msg.subject}", f"{i:05d}_mesaj") + ".eml"
            try:
                with open(os.path.join(dest, fname), "wb") as fh:
                    fh.write(message_to_eml_bytes(msg))
            except Exception as exc:
                report(f"  ! Yazilamadi: {fname} ({exc})")
            done += 1
            if done % 25 == 0 or done == total:
                report(f"{done}/{total} mesaj aktarildi", done / total)
        for sub in folder.subfolders:
            walk(sub, os.path.join(rel, _safe_name(sub.name, "klasor")))

    report("OST okunuyor...", 0.0)
    # Kok klasorun mesajlari out_dir'e, alt klasorler adlandirilmis alt
    # dizinlere yazilir (kok klasor adini tekrar etmeden).
    walk(tree, "")
    report("Bitti.", 1.0)
    return out_dir


def export_mbox(
    path: str,
    mbox_path: str,
    progress: Optional[Callable[[str, float], None]] = None,
) -> str:
    """Tum mesajlari tek bir MBOX dosyasina yazar."""
    import mailbox

    tree = read_ost(path)
    total = max(1, tree.total_messages())
    done = 0

    def report(m: str, f: float = -1.0) -> None:
        if progress:
            progress(m, f)

    mbox = mailbox.mbox(mbox_path)
    mbox.lock()
    try:
        def walk(folder: Folder) -> None:
            nonlocal done
            for msg in folder.messages:
                try:
                    mbox.add(message_to_eml_bytes(msg))
                except Exception:
                    pass
                done += 1
                if done % 25 == 0 or done == total:
                    report(f"{done}/{total} mesaj yazildi", done / total)
            for sub in folder.subfolders:
                walk(sub)

        report("OST okunuyor...", 0.0)
        walk(tree)
        mbox.flush()
    finally:
        mbox.unlock()
        mbox.close()
    report("Bitti.", 1.0)
    return mbox_path
