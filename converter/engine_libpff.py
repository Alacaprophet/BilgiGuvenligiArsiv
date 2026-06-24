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
    """OST/PST dosyasini okuyup klasor agacini dondurur."""
    if not is_available():
        raise RuntimeError("libpff (pypff) bulunamadi. 'pip install libpff-python'")
    import pypff

    if not os.path.exists(path):
        raise FileNotFoundError(path)

    pff = pypff.file()
    pff.open(path)
    try:
        root = pff.get_root_folder()
        tree = _read_folder(root)
        tree.name = os.path.splitext(os.path.basename(path))[0]
        return tree
    finally:
        pff.close()


# --------------------------------------------------------------------------- #
# EML uretimi
# --------------------------------------------------------------------------- #
_INVALID = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def _safe_name(name: str, fallback: str = "oge") -> str:
    name = _INVALID.sub("_", (name or "").strip()) or fallback
    return name[:120]


def message_to_eml_bytes(msg: Message) -> bytes:
    """Bir Message nesnesini RFC822 (.eml) baytlarina cevirir."""
    eml: EmailMessage

    if msg.headers.strip():
        # Orijinal basliklari koru, govdeyi yeniden ekle.
        parsed = Parser(policy=default_policy).parsestr(msg.headers, headersonly=True)
        eml = EmailMessage()
        for key, value in parsed.items():
            try:
                eml[key] = value
            except Exception:
                continue
    else:
        eml = EmailMessage()
        eml["Subject"] = msg.subject
        if msg.sender_name:
            eml["From"] = msg.sender_name

    if "Subject" not in eml:
        eml["Subject"] = msg.subject

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
