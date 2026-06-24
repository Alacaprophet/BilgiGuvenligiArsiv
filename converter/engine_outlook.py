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

    pst_path = os.path.abspath(pst_path)
    if not pst_path.lower().endswith(".pst"):
        pst_path += ".pst"
    if os.path.exists(pst_path):
        raise FileExistsError(
            f"Hedef dosya zaten var: {pst_path}\n"
            "Lutfen baska bir ad secin veya mevcut dosyayi tasiyin."
        )

    os.makedirs(os.path.dirname(pst_path) or ".", exist_ok=True)

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
    ns.AddStoreEx(pst_path, OL_STORE_UNICODE)

    # Yeni eklenen PST store'unu dosya yoluna gore bul.
    dest_store = None
    for i in range(1, ns.Stores.Count + 1):
        st = ns.Stores.Item(i)
        try:
            if (st.FilePath or "").lower() == pst_path.lower():
                dest_store = st
                break
        except Exception:
            continue
    if dest_store is None:
        raise RuntimeError("Olusturulan PST store profilde bulunamadi.")

    dest_root = dest_store.GetRootFolder()

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
    return pst_path


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

    libpff motoru orphan bir OST'yi .eml agacina cevirir; bu fonksiyon da onu
    Outlook'un olusturdugu GERCEK bir PST'ye yazar. Boylece sonuc dosyasi
    Outlook'ta sorunsuz acilir.
    """
    if not is_available():
        raise RuntimeError("Outlook / pywin32 bu makinede bulunamadi.")

    def report(msg: str, frac: float = -1.0) -> None:
        if progress:
            progress(msg, frac)

    pst_path = os.path.abspath(pst_path)
    if not pst_path.lower().endswith(".pst"):
        pst_path += ".pst"
    if os.path.exists(pst_path):
        raise FileExistsError(f"Hedef dosya zaten var: {pst_path}")
    os.makedirs(os.path.dirname(pst_path) or ".", exist_ok=True)

    ns = _namespace()
    report("Hedef PST dosyasi olusturuluyor...", 0.0)
    ns.AddStoreEx(pst_path, OL_STORE_UNICODE)

    dest_store = None
    for i in range(1, ns.Stores.Count + 1):
        st = ns.Stores.Item(i)
        try:
            if (st.FilePath or "").lower() == pst_path.lower():
                dest_store = st
                break
        except Exception:
            continue
    if dest_store is None:
        raise RuntimeError("Olusturulan PST store profilde bulunamadi.")
    dest_root = dest_store.GetRootFolder()

    total = max(1, _count_eml(eml_root))
    done = 0
    fail = 0
    fail_logged = 0
    # ~200 ilerleme guncellemesi -> akici arayuz, az COM/kuyruk yuku.
    step = max(25, total // 200)

    folder_cache = {}

    def get_or_create(parent, name: str):
        for i in range(1, parent.Folders.Count + 1):
            try:
                if parent.Folders.Item(i).Name == name:
                    return parent.Folders.Item(i)
            except Exception:
                continue
        return parent.Folders.Add(name)

    def folder_for(rel: str):
        """rel yolu icin (onbellekli) hedef klasoru dondurur."""
        if not rel or rel == ".":
            return dest_root
        if rel in folder_cache:
            return folder_cache[rel]
        folder = dest_root
        for part in rel.split(os.sep):
            folder = get_or_create(folder, part)
        folder_cache[rel] = folder
        return folder

    for current_dir, _dirs, files in os.walk(eml_root):
        rel = os.path.relpath(current_dir, eml_root)
        try:
            folder = folder_for(rel)
        except Exception as exc:
            report(f"  ! Klasor olusturulamadi ({rel}): {exc}")
            continue

        for fname in files:
            if not fname.lower().endswith(".eml"):
                continue
            full = os.path.join(current_dir, fname)
            item = None
            try:
                item = ns.OpenSharedItem(full)
                item.Move(folder)
            except Exception as exc:  # pragma: no cover - Outlook'a bagli
                fail += 1
                if fail_logged < 15:
                    fail_logged += 1
                    report(f"  ! {fname} aktarilamadi: {exc}")
                elif fail_logged == 15:
                    fail_logged += 1
                    report("  ! (daha fazla aktarilamayan oge sessizce gecilecek)")
            finally:
                item = None
            done += 1
            if done % step == 0 or done == total:
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
    return pst_path
