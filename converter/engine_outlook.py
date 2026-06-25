"""Microsoft Outlook (COM / MAPI) tabanli OST -> PST motoru.

Mantik:
    Outlook bir ``.ost`` dosyasini dogrudan "store" olarak ekleyemez (Microsoft
    buna izin vermez). Ancak OST dosyasi, profilde tanimli bir Exchange / IMAP
    hesabinin onbellegidir. Bu yuzden en saglam yontem sudur:

        1. Yeni, bos bir PST store olustur (``AddStoreEx``).
        2. Kaynak posta kutusunun (OST onbellegi) tum klasorlerini bu yeni
           PST'ye kopyala (``MAPIFolder.CopyTo``).
        3. PST store'unu profilden cikar (``RemoveStore``) -> dosya kapanir ve
           tasinabilir hale gelir.

    Sonucta olusan PST dosyasini Outlook bizzat olusturdugu icin baska bir
    Outlook kurulumunda da sorunsuz acilir.

Bu modul yalnizca Windows + Outlook + pywin32 ortaminda calisir. Diger
platformlarda ``is_available()`` False doner ve motor sessizce devre disi kalir.
"""

from __future__ import annotations

import functools
import os
import sys
from typing import Callable, List, Optional

#: engine_libpff'in her klasor dizinine yazdigi, klasorun GERCEK adini tutan
#: yan dosya. Aktarimda Outlook klasoru bu adla olusturulur (birebir yapi).
FOLDERNAME_FILE = "__foldername__.txt"


def _with_com(fn):
    """COM cagrilarini bir is parcaciginda guvenli kilar.

    Donusumler arayuzu dondurmamak icin ayri bir is parcaciginda calisir.
    Outlook COM nesneleri o parcacikta ``CoInitialize`` gerektirir; aksi halde
    "CoInitialize has not been called" hatasi olusur. Bu sarmalayici her cagri
    icin COM'u baslatip kapatir (ana parcacikta da zararsizdir).
    """

    @functools.wraps(fn)
    def wrapper(*args, **kwargs):
        try:
            import pythoncom
        except Exception:
            return fn(*args, **kwargs)
        pythoncom.CoInitialize()
        try:
            return fn(*args, **kwargs)
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    return wrapper

# OlStoreType sabitleri (Outlook Object Model).
# Unicode store 2 GB ANSI sinirini ortadan kaldirir -> buyuk posta kutulari icin.
OL_STORE_DEFAULT = 1
OL_STORE_UNICODE = 3
OL_STORE_ANSI = 4

# Kopyalanirken atlanacak ozel / kopyalanamayan klasorler.
_SKIP_FOLDERS = {
    "Sync Issues",
    "Senkronizasyon Sorunlari",
    "Conversation Action Settings",
    "Quick Step Settings",
    "RSS Feeds",
    "RSS Aboneleri",
}


def is_available() -> bool:
    """pywin32 ve Outlook COM bu makinede kullanilabilir mi?"""
    if not sys.platform.startswith("win"):
        return False
    try:
        import win32com.client  # noqa: F401
    except Exception:
        return False
    return True


class StoreInfo:
    """Outlook profilinde tanimli bir posta deposu (store) hakkinda bilgi."""

    def __init__(self, store_id: str, name: str, file_path: str):
        self.store_id = store_id
        self.name = name
        self.file_path = file_path or ""

    @property
    def is_ost(self) -> bool:
        return self.file_path.lower().endswith(".ost")

    @property
    def kind(self) -> str:
        ext = os.path.splitext(self.file_path)[1].lower()
        return {".ost": "OST", ".pst": "PST"}.get(ext, "Exchange/Diger")

    def __str__(self) -> str:
        return f"{self.name}  [{self.kind}]"


def _namespace():
    import win32com.client

    outlook = win32com.client.Dispatch("Outlook.Application")
    return outlook.GetNamespace("MAPI")


@_with_com
def list_stores() -> List[StoreInfo]:
    """Outlook profilindeki tum store'lari dondurur.

    Kullanici, OST onbellegine sahip Exchange/IMAP hesabini bu listeden secer.
    """
    if not is_available():
        return []

    ns = _namespace()
    stores: List[StoreInfo] = []
    for i in range(1, ns.Stores.Count + 1):
        try:
            store = ns.Stores.Item(i)
            file_path = ""
            try:
                file_path = store.FilePath or ""
            except Exception:
                file_path = ""
            stores.append(StoreInfo(store.StoreID, store.DisplayName, file_path))
        except Exception:
            # Erisilemeyen store'u atla.
            continue
    return stores


def _count_items(folder) -> int:
    """Bir klasor ve alt klasorlerindeki toplam oge sayisi (ilerleme icin)."""
    total = 0
    try:
        total += folder.Items.Count
    except Exception:
        pass
    try:
        for i in range(1, folder.Folders.Count + 1):
            total += _count_items(folder.Folders.Item(i))
    except Exception:
        pass
    return total


def _prepare_pst_path(pst_path: str) -> str:
    """PST yolunu yerel (ters egik cizgili) mutlak bicime getirir ve dogrular."""
    pst_path = os.path.normpath(os.path.abspath(pst_path))
    if not pst_path.lower().endswith(".pst"):
        pst_path += ".pst"
    if os.path.exists(pst_path):
        raise FileExistsError(
            f"Hedef dosya zaten var: {pst_path}\n"
            "Lutfen baska bir ad secin veya mevcut dosyayi tasiyin."
        )
    os.makedirs(os.path.dirname(pst_path) or ".", exist_ok=True)
    return pst_path


def _remove_ghost_pst_stores(ns, report=None) -> int:
    """Dosyasi diskte bulunmayan (hayalet) PST store'larini profilden cikarir.

    Bunlar onceki basarisiz/yarim kalan denemelerden Outlook profiline takili
    kalan girislerdir; Outlook acilista veya AddStoreEx sirasinda bunlar icin
    'dosya bulunamiyor' uyarisi verip islemi kilitler. Temizleyince akis acilir.
    """
    removed = 0
    try:
        count = ns.Stores.Count
    except Exception:
        return 0
    # Geriye dogru gez: store cikarilinca sonraki indeksler kayar.
    for i in range(count, 0, -1):
        try:
            st = ns.Stores.Item(i)
            fp = st.FilePath or ""
        except Exception:
            continue
        if not fp.lower().endswith(".pst"):
            continue
        try:
            if os.path.exists(fp):
                continue  # gecerli PST; dokunma
        except Exception:
            continue
        try:
            ns.RemoveStore(st.GetRootFolder())
            removed += 1
        except Exception:
            pass
    if removed and report:
        report(f"Onceki denemelerden kalan {removed} eksik PST kaydi temizlendi.")
    return removed


def _create_pst_store(ns, pst_path: str, report=None):
    """Yeni bir Unicode PST store olusturup dondurur.

    Outlook bazen (kurumsal politika veya yol kisitlari yuzunden) PST'yi
    istenen klasor yerine varsayilan konuma (orn. Documents) koyar. Bu yuzden
    yeni store'u YOLA gore degil, "az once eklenen store" olarak yakalariz;
    boylece dosya nereye olusturulursa olusturulsun donusume devam edebiliriz.

    Donus: (dest_store, gercek_pst_yolu)
    """
    # Onceki denemelerden kalan hayalet PST kayitlarini temizle; aksi halde
    # Outlook 'dosya bulunamiyor' uyarisi verip AddStoreEx'i kilitleyebilir.
    _remove_ghost_pst_stores(ns, report)

    before = set()
    for i in range(1, ns.Stores.Count + 1):
        try:
            before.add(ns.Stores.Item(i).StoreID)
        except Exception:
            continue

    try:
        ns.AddStoreEx(pst_path, OL_STORE_UNICODE)
    except Exception as exc:
        raise RuntimeError(
            "PST olusturulamadi / bu oturumda yuklenemedi: %s\n\n"
            "EN OLASI SEBEP: Bu makinedeki Outlook'ta tanimli bir PROFIL yok. "
            "Outlook, profil olmadan PST veri dosyasi yukleyemez (bu yuzden "
            "mesajlar PST'ye yazilmaz, dosya bos/265 KB kalir).\n\n"
            "COZUM 1 (PST icin): Outlook'a bir profil tanimlayin -> Denetim "
            "Masasi > 'Mail (Microsoft Outlook)' > 'Profilleri Goster' > 'Ekle' "
            "ile bir profil olusturun (herhangi bir hesap olabilir; salih "
            "hesabinin bilgileri SART DEGIL), Outlook'u bir kez acin, sonra "
            "tekrar deneyin.\n\n"
            "COZUM 2 (profilsiz, garanti): 'Cikti bicimi = EML klasoru' secin; "
            "Outlook gerektirmez, tum icerigi eksiksiz verir." % exc
        ) from exc

    dest_store = None
    # Once tam yol eslesmesini dene (istenen konuma olusturulduysa).
    for i in range(1, ns.Stores.Count + 1):
        st = ns.Stores.Item(i)
        try:
            if (st.FilePath or "").lower() == pst_path.lower():
                dest_store = st
                break
        except Exception:
            continue
    # Bulunamazsa: yeni eklenen store hangisiyse onu al (Outlook baska yere
    # koymus olabilir).
    if dest_store is None:
        for i in range(1, ns.Stores.Count + 1):
            st = ns.Stores.Item(i)
            try:
                if st.StoreID not in before:
                    dest_store = st
                    break
            except Exception:
                continue
    if dest_store is None:
        raise RuntimeError(
            "PST dosyasi olusturulamadi: Outlook yeni veri dosyasini eklemedi.\n\n"
            "Bu genellikle kurumsal bir guvenlik politikasinin (Group Policy) yeni "
            "PST olusturmayi engellemesinden kaynaklanir.\n"
            "Cozum: 'Cikti bicimi' olarak EML klasoru veya MBOX secin "
            "(Outlook gerektirmez, ayni icerigi kurtarir)."
        )
    try:
        actual = dest_store.FilePath or pst_path
    except Exception:
        actual = pst_path
    return dest_store, actual


@_with_com
def convert_store_to_pst(
    store_id: str,
    pst_path: str,
    progress: Optional[Callable[[str, float], None]] = None,
) -> str:
    """Secilen store'u (OST onbellegi) yeni bir PST dosyasina kopyalar.

    Args:
        store_id: ``list_stores()`` ile elde edilen kaynak store kimligi.
        pst_path: Olusturulacak ``.pst`` dosyasinin tam yolu.
        progress: ``(mesaj, oran)`` cagrilan ilerleme geri cagrimi (oran 0..1).

    Returns:
        Olusturulan PST dosyasinin yolu.
    """
    if not is_available():
        raise RuntimeError("Outlook / pywin32 bu makinede bulunamadi.")

    def report(msg: str, frac: float = -1.0) -> None:
        if progress:
            progress(msg, frac)

    pst_path = _prepare_pst_path(pst_path)

    ns = _namespace()

    # Kaynak store'u bul.
    source_store = None
    for i in range(1, ns.Stores.Count + 1):
        st = ns.Stores.Item(i)
        if st.StoreID == store_id:
            source_store = st
            break
    if source_store is None:
        raise ValueError("Kaynak posta kutusu (store) bulunamadi.")

    source_root = source_store.GetRootFolder()

    report("Hedef PST dosyasi olusturuluyor...", 0.0)
    dest_store, actual_path = _create_pst_store(ns, pst_path, report)
    dest_root = dest_store.GetRootFolder()
    if actual_path and actual_path.lower() != pst_path.lower():
        report(f"Not: PST su konumda olusturuldu: {actual_path}")

    # Kopyalanacak ust duzey klasorleri topla.
    top_folders = []
    for i in range(1, source_root.Folders.Count + 1):
        f = source_root.Folders.Item(i)
        if f.Name in _SKIP_FOLDERS:
            continue
        top_folders.append(f)

    report("Ogeler sayiliyor...", 0.02)
    grand_total = max(1, sum(_count_items(f) for f in top_folders))

    copied = 0
    for idx, folder in enumerate(top_folders, start=1):
        name = folder.Name
        report(f"Kopyalaniyor: {name}", copied / grand_total)
        try:
            folder.CopyTo(dest_root)
        except Exception as exc:  # pragma: no cover - Outlook'a bagli
            report(f"  ! {name} kopyalanamadi: {exc}", -1.0)
        copied += _count_items(folder)
        report(f"Tamamlandi: {name}", min(0.99, copied / grand_total))

    # PST'yi profilden cikar -> dosya serbest kalir ve tasinabilir olur.
    report("PST dosyasi kapatiliyor...", 0.99)
    try:
        ns.RemoveStore(dest_root)
    except Exception:
        # Cikarilamadiysa dosya yine de gecerlidir; sadece Outlook'ta acik kalir.
        pass

    report("Bitti.", 1.0)
    return actual_path


def _count_eml(root_dir: str) -> int:
    total = 0
    for _, _, files in os.walk(root_dir):
        total += sum(1 for f in files if f.lower().endswith(".eml"))
    return total


@_with_com
def import_eml_tree_to_pst(
    eml_root: str,
    pst_path: str,
    progress: Optional[Callable[[str, float], None]] = None,
) -> str:
    """Bir .eml klasor agacini Outlook araciligi ile yeni bir PST'ye aktarir.

    Once ``OpenSharedItem`` (Outlook'un yerlesik .eml ice aktarimi) denenir;
    ``.eml`` dosya iliskisi olmayan makinelerde ('Gecersiz yol/URL') ise ayni
    .eml AYRISTIRILIP mesaj DOGRUDAN olusturulur. Boylece her iki ortamda da
    calisir. Outlook cokerse yeniden baglanir, her 500 mesajda PST diske yazilir.
    """
    import re as _re
    import gc as _gc
    import time as _time
    import email as _email
    import tempfile as _tempfile
    from email import policy as _email_policy
    from email.utils import parsedate_to_datetime as _parsedate

    if not is_available():
        raise RuntimeError("Outlook / pywin32 bu makinede bulunamadi.")

    def report(msg: str, frac: float = -1.0) -> None:
        if progress:
            progress(msg, frac)

    pst_path = _prepare_pst_path(pst_path)
    total = max(1, _count_eml(eml_root))
    done = 0
    fail = 0
    fail_logged = 0
    step = max(25, total // 200)
    FLUSH_EVERY = 500
    used_fallback = {"v": False}

    state = {"ns": None, "root": None}
    folder_cache = {}

    def _connect(create):
        ns2 = _namespace()
        if create:
            store, actual = _create_pst_store(ns2, pst_path, report)
        else:
            store = _add_existing_pst_store(ns2, pst_path)
            try:
                actual = store.FilePath or pst_path
            except Exception:
                actual = pst_path
        state["ns"] = ns2
        state["root"] = store.GetRootFolder()
        folder_cache.clear()
        folder_cache["."] = state["root"]
        return actual

    def _reconnect():
        for attempt in range(6):
            try:
                _time.sleep(min(2 * (attempt + 1), 10))
                _connect(create=False)
                return True
            except Exception:
                continue
        return False

    def _flush_reopen():
        try:
            state["ns"].RemoveStore(state["root"])
        except Exception:
            pass
        _gc.collect()
        try:
            _connect(create=False)
            return True
        except Exception:
            return _reconnect()

    def _is_disconnect(exc):
        try:
            return bool(exc.args) and exc.args[0] in _DISCONNECT_CODES
        except Exception:
            return False

    def get_or_create(parent, name):
        for i in range(1, parent.Folders.Count + 1):
            try:
                if parent.Folders.Item(i).Name == name:
                    return parent.Folders.Item(i)
            except Exception:
                continue
        return parent.Folders.Add(name)

    def display_name_for(dir_path, fallback):
        side = os.path.join(dir_path, FOLDERNAME_FILE)
        try:
            if os.path.exists(side):
                with open(side, "r", encoding="utf-8") as fh:
                    name = fh.read().strip()
                    if name:
                        return name
        except Exception:
            pass
        return fallback

    def folder_for_rel(rel):
        if rel in folder_cache:
            return folder_cache[rel]
        if not rel or rel == ".":
            folder_cache["."] = state["root"]
            return state["root"]
        parent_rel = os.path.dirname(rel) or "."
        parent = folder_for_rel(parent_rel)
        dname = display_name_for(os.path.join(eml_root, rel), os.path.basename(rel))
        folder = get_or_create(parent, dname)
        folder_cache[rel] = folder
        return folder

    _bad = _re.compile(r'[<>:"/\\|?*\x00-\x1f]')
    att_dir = _tempfile.mkdtemp(prefix="ost2pst_att_")

    def _direct_from_eml(folder, path):
        """OpenSharedItem calismadiginda: .eml'i ayristirip mesaji dogrudan kur."""
        with open(path, "rb") as fh:
            em = _email.message_from_binary_file(fh, policy=_email_policy.default)
        item = folder.Items.Add("IPM.Note")
        item.Subject = (em.get("Subject") or "(konusuz)")
        plain = html = ""
        try:
            bp = em.get_body(preferencelist=("plain",))
            if bp is not None:
                plain = bp.get_content()
        except Exception:
            pass
        try:
            bh = em.get_body(preferencelist=("html",))
            if bh is not None:
                html = bh.get_content()
        except Exception:
            pass
        if html:
            try:
                item.HTMLBody = html
            except Exception:
                item.Body = plain or ""
        else:
            item.Body = plain or ""
        pa = item.PropertyAccessor

        def setp(tag, val):
            try:
                pa.SetProperty(tag, val)
            except Exception:
                pass

        frm = em.get("From", "") or ""
        to = em.get("To", "") or ""
        cc = em.get("Cc", "") or ""
        if frm:
            setp(_PR_SENDER_NAME, frm)
            setp(_PR_SENT_REPR_NAME, frm)
        if to:
            setp(_PR_DISPLAY_TO, to)
        if cc:
            setp(_PR_DISPLAY_CC, cc)
        dt = em.get("Date")
        if dt:
            try:
                d = _parsedate(dt)
                if d:
                    setp(_PR_DELIVERY_TIME, d)
                    setp(_PR_SUBMIT_TIME, d)
            except Exception:
                pass
        for idx, att in enumerate(em.iter_attachments()):
            try:
                data = att.get_payload(decode=True) or b""
                name = _bad.sub("_", (att.get_filename() or "ek_%d.bin" % idx))[:120]
                pth = os.path.join(att_dir, name or ("ek_%d.bin" % idx))
                with open(pth, "wb") as fh:
                    fh.write(data)
                item.Attachments.Add(pth)
                try:
                    os.remove(pth)
                except Exception:
                    pass
            except Exception:
                continue
        item.Save()

    def _write_one(rel, full):
        folder = folder_for_rel(rel)
        # 1) Outlook'un yerlesik .eml ice aktarimi (en yuksek sadakat).
        try:
            it = state["ns"].OpenSharedItem(full)
            it.Move(folder)
            return
        except Exception as exc:
            if _is_disconnect(exc):
                raise  # disti loop yeniden baglansin
            # .eml iliskisi yok / acamadi -> ayristirip dogrudan kur.
            used_fallback["v"] = True
        _direct_from_eml(folder, full)

    report("Hedef PST dosyasi olusturuluyor...", 0.0)
    actual_path = _connect(create=True)
    report("Aktariliyor...", 0.0)
    aborted = False
    try:
        for current_dir, _dirs, files in os.walk(eml_root):
            rel = os.path.relpath(current_dir, eml_root)
            for fname in files:
                if fname == FOLDERNAME_FILE or not fname.lower().endswith(".eml"):
                    continue
                full = os.path.join(current_dir, fname)
                ok = False
                try:
                    _write_one(rel, full)
                    ok = True
                except Exception as exc:
                    if _is_disconnect(exc):
                        report("  ! Outlook baglantisi koptu; yeniden "
                               "baglaniliyor...")
                        if _reconnect():
                            try:
                                _write_one(rel, full)
                                ok = True
                            except Exception:
                                ok = False
                        else:
                            report("  ! Outlook'a yeniden baglanilamadi; kalanlar "
                                   "atlandi.")
                            aborted = True
                            break
                    else:
                        ok = False
                if not ok:
                    fail += 1
                    if fail_logged < 15:
                        fail_logged += 1
                        try:
                            emsg = str(exc)
                        except Exception:
                            emsg = "?"
                        report(f"  ! {fname} aktarilamadi: {emsg}")
                    elif fail_logged == 15:
                        fail_logged += 1
                        report("  ! (daha fazla aktarilamayan oge gecilecek)")
                done += 1
                if done % step == 0 or done >= total:
                    report(f"{done}/{total} mesaj PST'ye yazildi",
                           min(1.0, done / total))
                if done % 250 == 0:
                    _gc.collect()
                if done % FLUSH_EVERY == 0:
                    report("  (ara kayit: PST diske yaziliyor...)")
                    if not _flush_reopen():
                        report("  ! Yeniden acilamadi; kalanlar atlandi.")
                        aborted = True
                        break
            if aborted:
                break
    finally:
        import shutil as _shutil
        _shutil.rmtree(att_dir, ignore_errors=True)

    report("PST dosyasi kapatiliyor...", 0.99)
    try:
        state["ns"].RemoveStore(state["root"])
    except Exception:
        pass

    success = done - fail
    if done > 0 and success == 0:
        try:
            if actual_path and os.path.exists(actual_path):
                os.remove(actual_path)
        except Exception:
            pass
        raise RuntimeError(
            "Hicbir mesaj PST'ye yazilamadi.\n"
            "Outlook bu makinede mesajlari PST'ye islemiyor olabilir "
            "(profilsiz/kararsiz Outlook). 'Cikti bicimi = EML klasoru' ile "
            "icerigi tam alabilir, PST'yi Outlook'un duzgun calistigi bir "
            "makinede uretebilirsiniz."
        )

    tail = " (yarida kesildi)" if aborted else ""
    if used_fallback["v"]:
        tail += " [.eml iliskisi yok -> dogrudan olusturma kullanildi]"
    if fail:
        report(f"Bitti. {success}/{total} mesaj yazildi, {fail} atlandi.{tail}", 1.0)
    else:
        report(f"Bitti. {success} mesaj.{tail}", 1.0)
    return actual_path




# MAPI proptag adlari (PropertyAccessor icin). .eml/OpenSharedItem KULLANMADAN
# mesaj olusturmak icin gerekli alanlar.
_PR_SENDER_NAME = "http://schemas.microsoft.com/mapi/proptag/0x0C1A001F"
_PR_SENDER_EMAIL = "http://schemas.microsoft.com/mapi/proptag/0x0C1F001F"
_PR_SENT_REPR_NAME = "http://schemas.microsoft.com/mapi/proptag/0x0042001F"
_PR_DELIVERY_TIME = "http://schemas.microsoft.com/mapi/proptag/0x0E060040"
_PR_SUBMIT_TIME = "http://schemas.microsoft.com/mapi/proptag/0x00390040"
_PR_DISPLAY_TO = "http://schemas.microsoft.com/mapi/proptag/0x0E04001F"
_PR_DISPLAY_CC = "http://schemas.microsoft.com/mapi/proptag/0x0E03001F"
_PR_HEADERS = "http://schemas.microsoft.com/mapi/proptag/0x007D001F"


def _add_existing_pst_store(ns, pst_path: str):
    """Var olan bir PST'yi (yeniden) acar ve store nesnesini dondurur."""
    before = set()
    for i in range(1, ns.Stores.Count + 1):
        try:
            before.add(ns.Stores.Item(i).StoreID)
        except Exception:
            continue
    ns.AddStoreEx(pst_path, OL_STORE_UNICODE)  # var olan dosyayi acar
    for i in range(1, ns.Stores.Count + 1):
        st = ns.Stores.Item(i)
        try:
            if (st.FilePath or "").lower() == pst_path.lower():
                return st
        except Exception:
            continue
    for i in range(1, ns.Stores.Count + 1):
        st = ns.Stores.Item(i)
        try:
            if st.StoreID not in before:
                return st
        except Exception:
            continue
    raise RuntimeError("PST yeniden acilamadi.")


# Outlook COM baglantisinin koptugunu gosteren HRESULT kodlari.
_DISCONNECT_CODES = {
    -2147417848,  # RPC_E_DISCONNECTED: cagrilan nesne istemcilerinden ayrildi
    -2147023174,  # RPC_S_SERVER_UNAVAILABLE: RPC sunucusu kullanilamiyor
    -2147023169,  # RPC_S_SERVER_TOO_BUSY / fatal
    -2147023170,  # RPC_S_CALL_FAILED benzeri
    -2146959355,  # CO_E_SERVER_EXEC_FAILURE: Outlook baslatilamadi
}


@_with_com
def import_messages_to_pst(
    pst_path: str,
    message_source,
    total: int,
    progress: Optional[Callable[[str, float], None]] = None,
) -> str:
    """libpff'ten okunan mesajlari Outlook'ta DOGRUDAN olusturup PST'ye yazar.

    ``.eml`` dosyasi ve ``OpenSharedItem`` KULLANMAZ; bu yuzden ``.eml`` dosya
    iliskisi olmayan makinelerde (Windows Server vb.) de calisir. Yalnizca
    Outlook'un kurulu olmasini gerektirir.

    Cok sayida oge (on binlerce) olustururken Outlook surecinin cokmesine karsi
    DAYANIKLIDIR: baglanti koparsa otomatik yeniden baglanir, PST'yi yeniden acar
    ve kaldigi yerden devam eder. Ayrica bellegi sismekten korumak icin periyodik
    olarak PST'yi diske yazip yeniden acar.

    ``message_source`` : ``(chain, Message)`` ureten yineleyici.
    """
    import re as _re
    import gc as _gc
    import time as _time
    import tempfile as _tempfile
    import datetime as _dt
    from email.parser import Parser as _Parser
    from email.policy import default as _default_policy

    if not is_available():
        raise RuntimeError("Outlook / pywin32 bu makinede bulunamadi.")

    def report(msg: str, frac: float = -1.0) -> None:
        if progress:
            progress(msg, frac)

    pst_path = _prepare_pst_path(pst_path)
    total = max(1, int(total))
    done = 0
    fail = 0
    fail_logged = 0
    step = max(25, total // 200)
    FLUSH_EVERY = 500  # bu kadar mesajda bir PST'yi diske yaz + yeniden ac

    state = {"ns": None, "root": None, "path": pst_path}
    folder_cache = {}

    def _connect(create):
        ns2 = _namespace()
        if create:
            store, actual = _create_pst_store(ns2, pst_path, report)
        else:
            store = _add_existing_pst_store(ns2, pst_path)
            try:
                actual = store.FilePath or pst_path
            except Exception:
                actual = pst_path
        state["ns"] = ns2
        state["root"] = store.GetRootFolder()
        state["path"] = actual
        folder_cache.clear()
        folder_cache[()] = state["root"]
        return actual

    def _reconnect():
        for attempt in range(6):
            try:
                _time.sleep(min(2 * (attempt + 1), 10))
                _connect(create=False)
                return True
            except Exception:
                continue
        return False

    def _flush_reopen():
        # PST'yi diske yaz (kapat) ve yeniden ac -> Outlook bellegi serbest kalir.
        try:
            state["ns"].RemoveStore(state["root"])
        except Exception:
            pass
        _gc.collect()
        try:
            _connect(create=False)
            return True
        except Exception:
            return _reconnect()

    def _is_disconnect(exc):
        try:
            return bool(exc.args) and exc.args[0] in _DISCONNECT_CODES
        except Exception:
            return False

    def get_or_create(parent, name):
        for i in range(1, parent.Folders.Count + 1):
            try:
                if parent.Folders.Item(i).Name == name:
                    return parent.Folders.Item(i)
            except Exception:
                continue
        return parent.Folders.Add(name)

    def folder_for(chain):
        if chain in folder_cache:
            return folder_cache[chain]
        parent = folder_for(chain[:-1])
        folder = get_or_create(parent, chain[-1])
        folder_cache[chain] = folder
        return folder

    _bad = _re.compile(r'[<>:"/\\|?*\x00-\x1f]')

    def _att_name(name, idx):
        name = _bad.sub("_", (name or "").strip()) or ("ek_%d.bin" % idx)
        return name[:120]

    att_dir = _tempfile.mkdtemp(prefix="ost2pst_att_")

    def _write_one(chain, msg):
        folder = folder_for(tuple(chain))
        item = folder.Items.Add("IPM.Note")
        item.Subject = getattr(msg, "subject", "") or "(konusuz)"

        html = getattr(msg, "html_body", "") or ""
        if html:
            try:
                item.HTMLBody = html
            except Exception:
                item.Body = getattr(msg, "plain_body", "") or ""
        else:
            item.Body = getattr(msg, "plain_body", "") or ""

        to = cc = from_disp = ""
        headers = getattr(msg, "headers", "") or ""
        if headers:
            try:
                p = _Parser(policy=_default_policy).parsestr(headers, headersonly=True)
                to = p.get("To", "") or ""
                cc = p.get("Cc", "") or ""
                from_disp = p.get("From", "") or ""
            except Exception:
                pass

        pa = item.PropertyAccessor

        def setp(tag, val):
            try:
                pa.SetProperty(tag, val)
            except Exception:
                pass

        sender = getattr(msg, "sender_name", "") or from_disp
        if sender:
            setp(_PR_SENDER_NAME, sender)
            setp(_PR_SENT_REPR_NAME, sender)
        if to:
            setp(_PR_DISPLAY_TO, to)
        if cc:
            setp(_PR_DISPLAY_CC, cc)
        if headers:
            setp(_PR_HEADERS, headers)
        d = getattr(msg, "date", None)
        if isinstance(d, _dt.datetime):
            setp(_PR_DELIVERY_TIME, d)
            setp(_PR_SUBMIT_TIME, d)

        for idx, att in enumerate(getattr(msg, "attachments", []) or []):
            try:
                pth = os.path.join(att_dir, _att_name(att.name, idx))
                with open(pth, "wb") as fh:
                    fh.write(att.data or b"")
                item.Attachments.Add(pth)
                try:
                    os.remove(pth)
                except Exception:
                    pass
            except Exception:
                continue

        item.Save()

    report("Hedef PST dosyasi olusturuluyor...", 0.0)
    actual_path = _connect(create=True)
    report("Aktariliyor...", 0.0)
    aborted = False
    try:
        for chain, msg in message_source:
            ok = False
            try:
                _write_one(chain, msg)
                ok = True
            except Exception as exc:
                if _is_disconnect(exc):
                    report("  ! Outlook baglantisi koptu; yeniden baglaniliyor...")
                    if _reconnect():
                        try:
                            _write_one(chain, msg)  # taze baglantida bir kez dene
                            ok = True
                        except Exception:
                            ok = False  # bu mesaj sorunlu; atla
                    else:
                        report("  ! Outlook'a yeniden baglanilamadi; kalan "
                               "mesajlar atlandi.")
                        aborted = True
                        break
                else:
                    ok = False
            if not ok:
                fail += 1
                if fail_logged < 15:
                    fail_logged += 1
                    try:
                        emsg = str(exc)
                    except Exception:
                        emsg = "?"
                    report(f"  ! Mesaj yazilamadi: {emsg}")
                elif fail_logged == 15:
                    fail_logged += 1
                    report("  ! (daha fazla yazilamayan oge sessizce gecilecek)")
            done += 1
            if done % step == 0 or done >= total:
                report(f"{done}/{total} mesaj PST'ye yazildi", min(1.0, done / total))
            if done % 250 == 0:
                _gc.collect()  # COM nesnelerini serbest birak (Outlook'u rahatlatir)
            if done % FLUSH_EVERY == 0:
                report("  (ara kayit: PST diske yaziliyor...)")
                if not _flush_reopen():
                    report("  ! Yeniden acilamadi; kalan mesajlar atlandi.")
                    aborted = True
                    break
    finally:
        import shutil as _shutil
        _shutil.rmtree(att_dir, ignore_errors=True)

    report("PST dosyasi kapatiliyor...", 0.99)
    try:
        state["ns"].RemoveStore(state["root"])
    except Exception:
        pass

    success = done - fail
    if done > 0 and success == 0:
        try:
            if actual_path and os.path.exists(actual_path):
                os.remove(actual_path)
        except Exception:
            pass
        raise RuntimeError(
            "Hicbir mesaj PST'ye yazilamadi. Outlook mesaj olusturmayi reddetti "
            "veya surekli cokuyor."
        )

    tail = " (islem yarida kesildi)" if aborted else ""
    if fail:
        report(f"Bitti. {success}/{total} mesaj yazildi, {fail} oge atlandi.{tail}",
               1.0)
    else:
        report(f"Bitti. {success} mesaj.{tail}", 1.0)
    return actual_path


@_with_com
def import_stream_to_pst(
    pst_path: str,
    source,
    total: int,
    progress: Optional[Callable[[str, float], None]] = None,
) -> str:
    """Akan (chain, eml_path) ogelerini yeni bir PST'ye aktarir (boru hatti).

    ``source`` : ``(name_chain, eml_path)`` ogeleri ureten bir yineleyici.
        ``name_chain`` kok->klasor gercek ad demeti; ``eml_path`` gecici .eml
        dosyasi. ``eml_path`` None ise oge okunamamistir (sayilir, atlanir).
        Aktarilan her .eml dosyasi islendikten sonra silinir (disk kontrollu).

    libpff uretici parcacigi mesajlari okuyup yazarken bu tuketici Outlook'a
    eszamanli aktarir; boylece toplam sure iki asamanin toplami yerine
    yavas olanina yaklasir.
    """
    if not is_available():
        raise RuntimeError("Outlook / pywin32 bu makinede bulunamadi.")

    def report(msg: str, frac: float = -1.0) -> None:
        if progress:
            progress(msg, frac)

    pst_path = _prepare_pst_path(pst_path)

    ns = _namespace()
    report("Hedef PST dosyasi olusturuluyor...", 0.0)
    dest_store, actual_path = _create_pst_store(ns, pst_path, report)
    dest_root = dest_store.GetRootFolder()
    if actual_path and actual_path.lower() != pst_path.lower():
        report(f"Not: PST su konumda olusturuldu: {actual_path}")

    total = max(1, int(total))
    done = 0
    fail = 0
    fail_logged = 0
    step = max(25, total // 200)

    def get_or_create(parent, name: str):
        for i in range(1, parent.Folders.Count + 1):
            try:
                if parent.Folders.Item(i).Name == name:
                    return parent.Folders.Item(i)
            except Exception:
                continue
        return parent.Folders.Add(name)

    folder_cache = {(): dest_root}

    def folder_for(chain):
        if chain in folder_cache:
            return folder_cache[chain]
        parent = folder_for(chain[:-1])
        folder = get_or_create(parent, chain[-1])
        folder_cache[chain] = folder
        return folder

    report("Aktariliyor...", 0.0)
    for chain, eml_path in source:
        if eml_path is None:
            done += 1
        else:
            try:
                folder = folder_for(tuple(chain))
            except Exception as exc:
                report(f"  ! Klasor olusturulamadi: {exc}")
                folder = dest_root
            item = None
            try:
                item = ns.OpenSharedItem(eml_path)
                item.Move(folder)
            except Exception as exc:  # pragma: no cover - Outlook'a bagli
                fail += 1
                if fail_logged < 15:
                    fail_logged += 1
                    report(f"  ! Oge aktarilamadi: {exc}")
                elif fail_logged == 15:
                    fail_logged += 1
                    report("  ! (daha fazla aktarilamayan oge sessizce gecilecek)")
            finally:
                item = None
                try:
                    os.remove(eml_path)
                except Exception:
                    pass
            done += 1
        if done % step == 0 or done >= total:
            report(f"{done}/{total} mesaj PST'ye yazildi", min(1.0, done / total))

    report("PST dosyasi kapatiliyor...", 0.99)
    try:
        ns.RemoveStore(dest_root)
    except Exception:
        pass
    if fail:
        report(f"Bitti. {total - fail}/{total} mesaj yazildi, {fail} oge atlandi.",
               1.0)
    else:
        report("Bitti.", 1.0)
    return actual_path
