# Ke hoach khac phuc bo sot doan Word

## Van de

Khi dich Word tu English sang Vietnamese, mot so doan van doi luc van con tieng Anh trong file ket qua. Luong hien tai da co nut "Quet bo sot", nhung:

- Viec quet bo sot dang la thao tac thu cong sau khi dich xong.
- Detector `find_missed()` chi bat ban dich rong hoac giong het text goc.
- Neu model tra ve ban dich nua voi, van con cum/cau tieng Anh, detector cu khong xem la bo sot.
- Neu ghi lai DOCX bi fail am tham o mot paragraph, translations dict co the da co ban dich nhung output van con text goc.

## Muc tieu

- Sau moi lan dich, app tu dong quet lai cac block con dau hieu chua dich.
- Neu con sot, app tu dong goi API dich tiep truoc khi tao file download.
- Nut "Quet bo sot" thu cong van giu lai, nhung phai dung detector moi.
- Batch translate cung duoc quet bu tu dong.
- Them smoke test de tranh regression.

## Huong xu ly

1. Mo rong `find_missed()`:
   - Van bat translation rong hoac trung text goc.
   - Bat them residual source language theo cap ngon ngu.
   - Voi English -> Vietnamese, bo qua standard/code/unit/acronym duoc phep giu nguyen, nhung bat cau/cum English ro rang nhu `shall be installed`, `landing door`, `control panel`.
   - Ho tro quet output blocks sau khi apply DOCX de phat hien truong hop apply khong ghi duoc vao file.

2. Them auto-rescan sau full translation:
   - Chay toi da `AUTO_RESCAN_PASSES` vong quet bu.
   - Moi vong chi dich cac block bi nghi con sot.
   - Cong token/cost vao summary cuoi.
   - Neu sau gioi han van con sot, log canh bao ro trong UI.

3. Them auto-rescan cho batch:
   - Moi file batch sau dich lan dau se tu quet va dich bu truoc khi tao output.

4. Cap nhat validation/log:
   - Sau khi apply DOCX, quet output blocks.
   - Neu output van con dau hieu source-language residue, them warning vao validation va log so luong.

5. Them tests:
   - Exact original bi xem la missed.
   - Ban dich nua voi con English bi xem la missed.
   - Standard/code/proper short tokens khong bi false positive.
   - Output blocks con text goc duoc phat hien du translations dict da co ban dich.

## Gioi han chap nhan

Detector la heuristic, khong the phan biet hoan hao moi ten rieng/brand/code. De giam false positive, scan chi kich hoat khi co dau hieu source-language ro rang va bo qua nhom token duoc phep giu nguyen.
