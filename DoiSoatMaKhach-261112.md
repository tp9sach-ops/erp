/** ============================================================
 *  9 SẠCH - DoiSoatMaKhach.gs  (FILE MỚI - thêm vào project Apps Script)
 *  ĐỐI SOÁT & SỬA MÃ KHÁCH HÀNG trong DataKhachSi cho ĐÚNG với file
 *  "[NEW CODE] DANH SÁCH KHÁCH HÀNG THÁNG 8/26" (cột New-code = mã chuẩn).
 *
 *  CÁCH DÙNG (chạy tay trong Apps Script Editor):
 *   1) doiSoatMa_XemTruoc()  -> CHỈ lập báo cáo ở sheet "KetQuaDoiSoatMa", KHÔNG sửa data.
 *   2) Kiểm tra báo cáo (các dòng CONFLICT / AMBIGUOUS / DUP_FILE phải xử lý tay).
 *   3) doiSoatMa_ApDung()    -> tự sao lưu DataKhachSi rồi sửa mã; đồng thời đổi Mã KH
 *                              cũ -> mới trong sheet TheoDoiHoaDon (nếu có).
 *
 *  QUY TẮC: file New Code là CHUẨN. Khớp khách theo thứ tự tin cậy:
 *   đang giữ đúng mã (và đúng người) > SĐT > CCCD/MST + tên > tên + địa chỉ > mã cũ KHS duy nhất.
 *   Khớp được: SỬA mã/xưởng/tỉnh/nhóm theo file, chỉ BỔ SUNG ô còn trống (không ghi đè tên/SĐT/ngày đặt/loại giá).
 *   Không khớp ai: THÊM dòng mới. Mơ hồ / đụng mã người khác: KHÔNG tự sửa, ghi vào báo cáo.
 *  Dữ liệu file nhúng bên dưới (DSM_FILE) lấy từ file xlsx ngày 05/10/2026 (207 dòng có mã).
 * ============================================================ */

const DSM_FILE = [
{"c": "ĐNCHGL", "id": "056095011286", "nm": "HỘ KINH DOANH NGUYỄN QUỐC VIỆT GL (Cửa hàng Gia Lai)", "ad": "51 Lê Lợi, Phường Pleiku, Tỉnh Gia Lai, Việt Nam", "gp": "CH", "nk": "", "nn": "", "fc": "ĐN", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "CTCHĐT-001", "id": "093098001899", "nm": "HỘ KINH DOANH BÁNH SẦU RIÊNG - ĐỒNG THÁP", "ad": "Số 10, đường Nguyễn Huệ, Phường Cao Lãnh, tỉnh Đồng Tháp", "gp": "CH", "nk": "", "nn": "", "fc": "CT", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "CTCHAG-001", "id": "093098001899", "nm": "HỘ KINH DOANH BÁNH SẦU RIÊNG - RẠCH GIÁ (Cửa hàng Rạch Giá)", "ad": "Số 87, đường Nguyễn Trung Trực, Phường Rạch Giá, An Giang", "gp": "CH", "nk": "", "nn": "", "fc": "CT", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "CTCHĐO-001", "id": "044097005868", "nm": "HỘ KINH DOANH BÁNH SẦU RIÊNG 9 SẠCH (Cửa hàng Biên Hòa-Tư Hoàng)", "ad": "163 Phan Trung, KP 2, Phường Tam Hiệp, tỉnh Đồng Nai, Việt Nam", "gp": "CH", "nk": "", "nn": "", "fc": "CT", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "CTCHSG-001", "id": "044097005868", "nm": "BÁNH SẦU RIÊNG 9 SẠCH THỦ ĐỨC (Tư Hoàng)", "ad": "29 Hoàng Diệu 2, KP 4, Phường Linh Xuân, Thành phố Hồ Chí Minh, Việt Nam", "gp": "CH", "nk": "", "nn": "", "fc": "CT", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "CTCHSG-002", "id": "044097005868", "nm": "BÁNH SẦU RIÊNG 9 SẠCH THỦ DẦU MỘT (Tư Hoàng)", "ad": "228 Phú Lợi, KP 5, phường Phú Lợi, Tp. Hồ Chí Minh, Việt Nam.", "gp": "CH", "nk": "", "nn": "", "fc": "CT", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "CTCHSG-003", "id": "093190007787", "nm": "BÁNH SẦU RIÊNG 9 SẠCH - CN 3 (CHỊ QUYÊN-BÌNH THẠNH)", "ad": "213/1 Nguyễn Gia Trí, Phường Thạnh Mỹ Tây, Thành phố Hồ Chí Minh, Việt Nam", "gp": "CH", "nk": "", "nn": "", "fc": "CT", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "CTCHSG-004", "id": "093190007787", "nm": "HỘ KINH DOANH BÁNH SẦU RIÊNG 9 SẠCH (CHỊ QUYÊN-GÒ VẤP)", "ad": "312 Lê Văn Thọ, Phường Thông Tây Hội, Thành phố Hồ Chí Minh, Việt Nam", "gp": "CH", "nk": "", "nn": "", "fc": "CT", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "CTCHSG-005", "id": "093190007787", "nm": "BÁNH SẦU RIÊNG 9 SẠCH - CN4 (CHỊ QUYÊN-PHÚ NHUẬN)", "ad": "337 Phan Đình Phùng, Phường Cầu Kiệu, Thành phố Hồ Chí Minh, Việt Nam", "gp": "CH", "nk": "", "nn": "", "fc": "CT", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "ĐNCHTA-001", "id": "093190007787", "nm": "BÁNH SẦU RIÊNG 9 SẠCH - CN5 (CHỊ QUYÊN-THÁI NGUYÊN)", "ad": "19A Minh Cầu, Phường Phan Đình Phùng, tỉnh Thái Nguyên, Việt Nam", "gp": "CH", "nk": "", "nn": "", "fc": "CT", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "ĐNCHQI-001", "id": "093190007787", "nm": "BÁNH SẦU RIÊNG 9 SẠCH - CN6 (CHỊ QUYÊN-HẠ LONG)", "ad": "20B Cao Xanh, Phường Cao Xanh, Tỉnh Quảng Ninh, Việt Nam", "gp": "CH", "nk": "", "nn": "", "fc": "ĐN", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "ĐNCHHP-001", "id": "093190007787", "nm": "BÁNH SẦU RIÊNG 9 SẠCH - CN7 (CHỊ QUYÊN-HẢI DƯƠNG)", "ad": "18A Phạm Ngũ Lão, Phường Lê Thanh Nghị, Thành phố Hải Phòng, Việt Nam", "gp": "CH", "nk": "", "nn": "", "fc": "ĐN", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "ĐNCHHP-002", "id": "031305009280", "nm": "Vũ Thị Thuý Hằng (CHỊ QUYÊN-HẢI PHÒNG)", "ad": "58 Hồ Sen, Phường Lê Chân, Thành phố Hải Phòng", "gp": "CH", "nk": "", "nn": "", "fc": "ĐN", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "ĐNCHBN-001", "id": "027196001895", "nm": "Ngô Thị Thu Ngân (CCCD số: 027196001895-Cửa hàng Bắc Ninh)", "ad": "Khu 4, phường Kinh Bắc, tỉnh Bắc Ninh, Việt Nam", "gp": "CH", "nk": "", "nn": "", "fc": "ĐN", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "ĐNCHHN-001", "id": "001172013398", "nm": "HỘ KINH DOANH BÁNH SẦU RIÊNG SẠCH SẠCH (Trang Cún-Cửa hàng Hà Đông)", "ad": "Số 388 Nguyễn Trãi, Phường Đại Mỗ, Thành phố Hà Nội, Việt Nam", "gp": "CH", "nk": "", "nn": "", "fc": "ĐN", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "ĐNCHHN-002", "id": "001059003138", "nm": "HỘ KINH DOANH BÁNH SẦU RIÊNG SẠCH (Trang Cún-Cửa hàng Cầu Giấy)", "ad": "Số 193 Hồ Tùng Mậu, Phường Từ Liêm, Thành phố Hà Nội, Việt Nam", "gp": "CH", "nk": "", "nn": "", "fc": "ĐN", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "ĐNCHTH-001", "id": "049194001430", "nm": "HỘ KINH DOANH LÊ THỊ DIỄM", "ad": "Số nhà 468A Đường Trần Phú, Phường Hạc Thành, Tỉnh Thanh Hóa, Việt Nam", "gp": "CH", "nk": "", "nn": "", "fc": "ĐN", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "ĐNCHHN-003", "id": "093093005909", "nm": "HỘ KINH DOANH BÁNH SẦU RIÊNG HỮU PHÚC", "ad": "36 phố Ô Chợ Dừa, Phường Ô Chợ Dừa, Thành Phố Hà Nội, Việt Nam", "gp": "CH", "nk": "", "nn": "", "fc": "ĐN", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "ĐNCHĐN-001", "id": "048080000776", "nm": "HỘ KINH DOANH LÊ HỮU CẢNH (Tú Cẩm-Đống Đa)", "ad": "238 Đống Đa, Phường Hải Châu, thành phố Đà Nẵng, Việt Nam", "gp": "CH", "nk": "", "nn": "", "fc": "ĐN", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "ĐNCHĐN-002", "id": "0402335126", "nm": "CÔNG TY TNHH THƯƠNG MẠI KIM NGÂN PHÚC", "ad": "40 Bàu Trảng 5, Phường Thanh Khê, Thành phố Đà Nẵng, Việt Nam", "gp": "CH", "nk": "", "nn": "", "fc": "ĐN", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "CTCHVL-001", "id": "086195001358", "nm": "HỘ KINH DOANH BÁNH SẦU RIÊNG 9 SẠCH", "ad": "Số 77/1 đường Trưng Nữ Vương, Phường Long Châu, Tỉnh Vĩnh Long, Việt Nam", "gp": "CH", "nk": "", "nn": "", "fc": "CT", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "CTCHAG-002", "id": "089182011983", "nm": "HỘ KINH DOANH TIỆM BÁNH SẦU RIÊNG 9 SẠCH (Cửa hàng Long Xuyên)", "ad": "290 Lý Thái Tổ nối dài, Phường Long Xuyên, Tỉnh An Giang, Việt Nam", "gp": "CH", "nk": "", "nn": "", "fc": "CT", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "CTCHCT-001", "id": "092083002851", "nm": "HỘ KINH DOANH SẠCH 99 (Nguyễn Minh Hoàng)", "ad": "Số 172 đường Xuân Thuỷ, phường An Bình, thành phố Cần Thơ.", "gp": "CH", "nk": "", "nn": "", "fc": "CT", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "CTSTWM-THAODIEN", "id": "0104918404", "nm": "CÔNG TY CỔ PHẦN DỊCH VỤ THƯƠNG MẠI TỔNG HỢP WINCOMMERCE (WINMART THẢO ĐIỀN)", "ad": "Số 23 Lê Duẩn, Phường Sài Gòn, Thành phố Hồ Chí Minh, Việt Nam", "gp": "ST", "nk": "", "nn": "WINMART THẢO ĐIỀN", "fc": "CT", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "ĐNHCGC-0006", "id": "001094043329", "nm": "BÙI THẾ LÝ (CCCD số: 001094043329-Út Mơ)", "ad": "Thôn Ngãi Cầu, Xã An Khánh, Thành phố Hà Nội", "gp": "HC", "nk": "Thế Lý - Út Mơ", "nn": "", "fc": "ĐN", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "CTSTLT-WESTLAKE", "id": "0304741634-015", "nm": "CÔNG TY CỔ PHẦN TRUNG TÂM THƯƠNG MẠI LOTTE VIỆT NAM - CHI NHÁNH TÂY HỒ", "ad": "Tầng hầm B1, Lotte Mall Hà Nội, Số 272 Võ Chí Công, Phường Tây Hồ, Thành phố Hà Nội, Việt Nam", "gp": "ST", "nk": "", "nn": "", "fc": "CT", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "CTSTLT-NHATRANG", "id": "0304741634-011", "nm": "CÔNG TY CỔ PHẦN TRUNG TÂM THƯƠNG MẠI LOTTE VIỆT NAM - CHI NHÁNH NHA TRANG", "ad": "Số 58 đường 23/10, Phường Tây Nha Trang, Tỉnh Khánh Hòa, Việt Nam", "gp": "ST", "nk": "", "nn": "CÔNG TY CỔ PHẦN TRUNG TÂM THƯƠNG MẠI LOTTE VIỆT NAM - CHI NHÁNH NHA TRANG", "fc": "CT", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "CTSTLT-HANOICT", "id": "0304741634-008", "nm": "CÔNG TY CỔ PHẦN TRUNG TÂM THƯƠNG MẠI LOTTE VIỆT NAM - CHI NHÁNH BA ĐÌNH", "ad": "Tầng hầm 1 (B1), Trung tâm Lotte Hà Nội, số 54, đường Liễu Giai, Phường Giảng Võ, Thành phố Hà Nội, Việt Nam", "gp": "ST", "nk": "NTGold", "nn": "", "fc": "CT", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "ĐNHCTM-0002", "id": "033185010825", "nm": "Vũ Thị Hằng (CCCD số: 033185010825)", "ad": "Tổ 16A2, khu phố Quang Trung 5A, Phường Uông Bí, Tỉnh Quảng Ninh", "gp": "HC", "nk": "HỘ KINH DOANH VŨ THỊ HẰNG 1985 (Chị Hằng-v/c anh Lý)", "nn": "", "fc": "ĐN", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "ĐNHCGC-0004", "id": "001193003953", "nm": "HỘ KINH DOANH : HÀ THU HOÀI", "ad": "Số nhà 05, ngách 03, ngõ 130, đường Hồng Phong, thôn Ngãi Cầu, Xã An Khánh, Thành phố Hà Nội, Việt Nam", "gp": "HC", "nk": "HỘ KINH DOANH : HÀ THU HOÀI (Tiến Tú - Út Mơ)", "nn": "", "fc": "ĐN", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "ĐNHCGC-0008", "id": "036178022676", "nm": "VŨ THỊ THỦY (CCCD số: 036178022676) (Thu Hoàng - Hà Nội)", "ad": "Xóm Năm, Xã Hải An, Tỉnh Ninh Bình", "gp": "HC", "nk": "Thu Hoàng - Hà Nội", "nn": "", "fc": "ĐN", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "ĐNHCGC-0007", "id": "001175007158", "nm": "Nguyễn Thị Ất", "ad": "Thôn Ngãi Cầu, Xã An Khánh, Thành phố Hà Nội", "gp": "HC", "nk": "Nguyễn Thị Ất CCCD số: 001175007158 (Cô Ất - Út Mơ )", "nn": "", "fc": "ĐN", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "ĐNHCGC-0005", "id": "001182033834", "nm": "Đỗ Thị Hà (CCCD số: 001182033834)", "ad": "Thôn Ngãi Cầu, Xã An Khánh, Thành phố Hà Nội", "gp": "HC", "nk": "Hà - Út Mơ", "nn": "", "fc": "ĐN", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "ĐNHCGC-0009", "id": "KHS0009", "nm": "Bán cho người tiêu dùng-Vũ Đình Tùng", "ad": "13/137 Phan Bội Châu, khu phố Trung Mới, phường Sầm Sơn, tỉnh Thanh Hóa", "gp": "HC", "nk": "Dương Thị Hằng (Cô Hằng)", "nn": "", "fc": "ĐN", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "ĐNHCGC-0001", "id": "1201699862", "nm": "CÔNG TY TNHH KINH DOANH VÀ PHÁT TRIỂN SẦU RIÊNG ÚT MƠ", "ad": "01 Đường Muông Tư, Ấp Thuận Hoà, Phường Long Thuận, Tỉnh Đồng Tháp, Việt Nam", "gp": "HC", "nk": "Tình - Út Mơ", "nn": "", "fc": "ĐN", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "ĐNHCTM-0001", "id": "027080009234", "nm": "Lưu Đình Lý (CCCD số: 027080009234)", "ad": "Tổ 16 a2, Khu phố Quang Trung 5A, phường Uông Bí, tỉnh Quảng Ninh", "gp": "HC", "nk": "HỘ KINH DOANH LƯU ĐINH LÝ (Lưu Đình Lý - CCCD số: 027080009234)", "nn": "", "fc": "ĐN", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "ĐNHCGC-0003", "id": "001196034438", "nm": "HỘ KINH DOANH NGUYỄN THỊ HỒNG 123", "ad": "Số 1281 đường Chùa Tổng, thôn Ngãi Cầu, Xã An Khánh, Thành phố Hà Nội, Việt Nam", "gp": "HC", "nk": "HKD NGUYỄN THỊ HỒNG 123 (Nguyên Anh - Út Mơ)", "nn": "", "fc": "ĐN", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "ĐNHCGC-0002", "id": "082196007577", "nm": "Hồ Thị Huỳnh Mơ (CCCD số: 082196007577)", "ad": "01 Đường Muông Tư, Ấp Thuận Hòa, Phường Long Thuận, Tỉnh Đồng Tháp", "gp": "HC", "nk": "Mơ - Út Mơ", "nn": "", "fc": "ĐN", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "CTSTLT-VUNGTAU", "id": "0304741634-005", "nm": "CÔNG TY CỔ PHẦN TRUNG TÂM THƯƠNG MẠI LOTTE VIỆT NAM - CHI NHÁNH BÀ RỊA VŨNG TÀU", "ad": "Góc đường 3 tháng 2 và đường Thi Sách, Phường Tam Thắng, Thành phố Hồ Chí Minh, Việt Nam", "gp": "ST", "nk": "", "nn": "", "fc": "CT", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "CTSTLT-NTGOLD", "id": "0304741634-011", "nm": "CÔNG TY CỔ PHẦN TRUNG TÂM THƯƠNG MẠI LOTTE VIỆT NAM - CHI NHÁNH NHA TRANG", "ad": "Số 58 đường 23/10, Phường Tây Nha Trang, Tỉnh Khánh Hòa, Việt Nam", "gp": "ST", "nk": "", "nn": "CÔNG TY CỔ PHẦN TRUNG TÂM THƯƠNG MẠI LOTTE VIỆT NAM - CHI NHÁNH NHA TRANG GOLD COAST", "fc": "CT", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "ĐNSTTC-0001", "id": "0106188175", "nm": "CÔNG TY CỔ PHẦN ĐẠT PHÁT HÀ NỘI", "ad": "Số 40, Phố Vũ Xuân Thiều, Phường Phúc Lợi, Thành phố Hà Nội, Việt Nam", "gp": "ST", "nk": "Thành Công - Lào Cai", "nn": "", "fc": "ĐN", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "ĐNSTTC-0002", "id": "0106188175", "nm": "CÔNG TY CỔ PHẦN ĐẠT PHÁT HÀ NỘI", "ad": "Số 40, Phố Vũ Xuân Thiều, Phường Phúc Lợi, Thành phố Hà Nội, Việt Nam", "gp": "ST", "nk": "Thành Công - Sapa", "nn": "", "fc": "ĐN", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "CTSTLT-PHANTHIET", "id": "0304741634-002", "nm": "CÔNG TY CỔ PHẦN TRUNG TÂM THƯƠNG MẠI LOTTE VIỆT NAM - CHI NHÁNH BÌNH THUẬN", "ad": "Khu dân cư Hùng Vương I, Phường Phú Thủy, tỉnh Lâm Đồng, Việt Nam", "gp": "ST", "nk": "", "nn": "", "fc": "CT", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "ĐNSTWM-DANANG", "id": "0104918404-009", "nm": "CHI NHÁNH ĐÀ NẴNG - CÔNG TY CỔ PHẦN DỊCH VỤ THƯƠNG MẠI TỔNG HỢP WINCOMMERCE", "ad": "L2 -01 Tầng 2, TTTM Vincom Plaza, 910A Ngô Quyền, Phường An Hải, Thành phố Đà Nẵng, Việt Nam", "gp": "ST", "nk": "", "nn": "", "fc": "ĐN", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "CTSTWM-LANDMARK", "id": "0104918404", "nm": "CÔNG TY CỔ PHẦN DỊCH VỤ THƯƠNG MẠI TỔNG HỢP WINCOMMERCE (WINMART LANDMARK 81)", "ad": "Số 23 Lê Duẩn, Phường Sài Gòn, Thành phố Hồ Chí Minh, Việt Nam", "gp": "ST", "nk": "", "nn": "WINMART LANDMARK 81", "fc": "CT", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "ĐNHCTM-0003", "id": "027080009234", "nm": "HỘ KINH DOANH LƯU ĐINH LÝ (Lưu Đình Lý)", "ad": "Số nhà 6B, Tổ 16A2, khu phố Quang Trung 5A, Phường Uông Bí, Tỉnh Quảng Ninh, Việt Nam", "gp": "HC", "nk": "HỘ KINH DOANH LƯU ĐINH LÝ (Lưu Đình Lý - CCCD số: 027080009234)", "nn": "", "fc": "ĐN", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "ĐNHCTM-0004", "id": "033185010825", "nm": "HỘ KINH DOANH VŨ THỊ HẰNG 1985 (Vũ Thị Hằng)", "ad": "Tổ 16A2, khu phố Quang Trung 5A, Phường Uông Bí, Tỉnh Quảng Ninh, Việt Nam", "gp": "HC", "nk": "HỘ KINH DOANH VŨ THỊ HẰNG 1985 (Chị Hằng-v/c anh Lý)", "nn": "", "fc": "ĐN", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "CTSTLT-CANTHO", "id": "0304741634-007", "nm": "CÔNG TY CỔ PHẦN TRUNG TÂM THƯƠNG MẠI LOTTE VIỆT NAM - CHI NHÁNH CẦN THƠ", "ad": "84, Mậu Thân, Phường Cái Khế, TP. Cần Thơ, Việt Nam", "gp": "ST", "nk": "", "nn": "", "fc": "CT", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "ĐNSTST-0001", "id": "0300100037-028", "nm": "CHI NHÁNH TỔNG CÔNG TY THƯƠNG MẠI SÀI GÒN - TNHH MỘT THÀNH VIÊN - TRUNG TÂM THƯƠNG MẠI SATRA VÕ VĂN KIỆT", "ad": "Số 1466, đường Võ Văn Kiệt, Phường Bình Tiên, Thành phố Hồ Chí Minh, Việt Nam", "gp": "ST", "nk": "", "nn": "", "fc": "CT", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "CTSTLT-VINH", "id": "0304741634-013", "nm": "CÔNG TY CỔ PHẦN TRUNG TÂM THƯƠNG MẠI LOTTE VIỆT NAM - CHI NHÁNH VINH", "ad": "Đại lộ V.I.Lenin, Khối Yên Sơn, Phường Vinh Phú, Tỉnh Nghệ An, Việt Nam", "gp": "ST", "nk": "", "nn": "", "fc": "CT", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "CTSTWM-HUNGGIA", "id": "0104918404", "nm": "CÔNG TY CỔ PHẦN DỊCH VỤ THƯƠNG MẠI TỔNG HỢP WINCOMMERCE (WINMART HƯNG GIA)", "ad": "Số 23 Lê Duẩn, Phường Sài Gòn, Thành phố Hồ Chí Minh, Việt Nam", "gp": "ST", "nk": "", "nn": "WINMART HƯNG GIA", "fc": "CT", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "CTSTLT-NAMSG", "id": "0304741634", "nm": "CÔNG TY CỔ PHẦN TRUNG TÂM THƯƠNG MẠI LOTTE VIỆT NAM", "ad": "Số 469, đường Nguyễn Hữu Thọ, Phường Tân Hưng, Thành Phố Hồ Chí Minh, Việt Nam", "gp": "ST", "nk": "", "nn": "", "fc": "CT", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "CTSTLT-GOVAP", "id": "0304741634-010", "nm": "CÔNG TY CỔ PHẦN TRUNG TÂM THƯƠNG MẠI LOTTE VIỆT NAM - CHI NHÁNH GÒ VẤP", "ad": "Số 18, Đường Phan Văn Trị, Phường Gò Vấp, Thành phố Hồ Chí Minh, Việt Nam", "gp": "ST", "nk": "", "nn": "", "fc": "CT", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "CTSTLT-TANBINH", "id": "0304741634-006", "nm": "CÔNG TY CỔ PHẦN TRUNG TÂM THƯƠNG MẠI LOTTE VIỆT NAM - CHI NHÁNH TÂN BÌNH", "ad": "Số 20, đường Cộng Hòa, Phường Bảy Hiền, Thành phố Hồ Chí Minh, Việt Nam", "gp": "ST", "nk": "", "nn": "", "fc": "CT", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "CTSTLT-BINHDUONG", "id": "0304741634-003", "nm": "CÔNG TY CỔ PHẦN TRUNG TÂM THƯƠNG MẠI LOTTE VIỆT NAM - CHI NHÁNH BÌNH DƯƠNG", "ad": "Khu đô thị The Seasons Bình Dương, Phường Lái Thiêu, TP. Hồ Chí Minh, Việt Nam", "gp": "ST", "nk": "", "nn": "", "fc": "CT", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "ĐNSTLT-DANANG", "id": "0304741634-009", "nm": "CÔNG TY CỔ PHẦN TRUNG TÂM THƯƠNG MẠI LOTTE VIỆT NAM - CHI NHÁNH ĐÀ NẴNG", "ad": "Số 06 Đường Nại Nam, Phường Hòa Cường, Thành phố Đà Nẵng, Việt Nam", "gp": "ST", "nk": "", "nn": "", "fc": "ĐN", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "ĐNHCGC-0010", "id": "038090041659", "nm": "Dương Thị Hằng (CCCD số: 038194004891)", "ad": "13/137 Phan Bội Châu, khu phố Trung Mới, phường Sầm Sơn, tỉnh Thanh Hóa", "gp": "HC", "nk": "Dương Thị Hằng (Cô Hằng)", "nn": "", "fc": "ĐN", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "ĐNHCGC-0011", "id": "075091001065", "nm": "HỘ KINH DOANH CỬA HÀNG ADAM BÁNH SẦU RIÊNG 9 SẠCH (CCCD: 075091001065)", "ad": "Ấp Hữu Thời, Xã Biển Bạch, Tỉnh Cà Mau, Việt Nam", "gp": "HC", "nk": "Anh Hiếu - Adam", "nn": "", "fc": "ĐN", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "ĐNHCTM-0005", "id": "033088013267", "nm": "HỘ KINH DOANH VŨ GIA HUY", "ad": "Thôn Dịch Trì, Xã Nguyễn Văn Linh, Tỉnh Hưng Yên, Việt Nam", "gp": "HC", "nk": "Vũ Gia Huy (V/c C Hằng ,a Lý )", "nn": "", "fc": "ĐN", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "ĐNSTWM-DANANG", "id": "0104918404-009", "nm": "CHI NHÁNH ĐÀ NẴNG - CÔNG TY CỔ PHẦN DỊCH VỤ THƯƠNG MẠI TỔNG HỢP WINCOMMERCE", "ad": "L2 -01 Tầng 2, TTTM Vincom Plaza, 910A Ngô Quyền, Phường An Hải, Thành phố Đà Nẵng, Việt Nam", "gp": "ST", "nk": "", "nn": "WINMART ĐÀ NẴNG", "fc": "ĐN", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "CTSTWM-DONGKHOI", "id": "0104918404", "nm": "CÔNG TY CỔ PHẦN DỊCH VỤ THƯƠNG MẠI TỔNG HỢP WINCOMMERCE (WINMART ĐỒNG KHỞI)", "ad": "Số 23 Lê Duẩn, Phường Sài Gòn, Thành phố Hồ Chí Minh, Việt Nam", "gp": "ST", "nk": "", "nn": "WINMART ĐỒNG KHỞI", "fc": "CT", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "CTSTWM-DIAMOND", "id": "0104918404", "nm": "CÔNG TY CỔ PHẦN DỊCH VỤ THƯƠNG MẠI TỔNG HỢP WINCOMMERCE (WINMART DIAMOND)", "ad": "Số 23 Lê Duẩn, Phường Sài Gòn, Thành phố Hồ Chí Minh, Việt Nam", "gp": "ST", "nk": "", "nn": "WINMART DIAMOND", "fc": "CT", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "CTSTLT-BACGIANG", "id": "0304741634-016", "nm": "CÔNG TY CỔ PHẦN TRUNG TÂM THƯƠNG MẠI LOTTE VIỆT NAM - CHI NHÁNH BẮC GIANG", "ad": "Lô TM Trung tâm VHTTDV TP Bắc Giang, Đường Hoàng Văn Thụ, Phường Bắc Giang, TP Bắc Ninh, Việt Nam", "gp": "ST", "nk": "", "nn": "LOTTE - BẮC GIANG", "fc": "CT", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "CTCHCT-002", "id": "091202009491", "nm": "HỘ KINH DOANH PHẠM KIM ĐỒNG", "ad": "18, đường Xô Viết Nghệ Tĩnh, Khóm 3, Phường Phú Lợi, Thành phố Cần Thơ", "gp": "CH", "nk": "DONGDOLA", "nn": "", "fc": "CT", "pv": "", "sh": "CH-ST-HC-XK"},
{"c": "CTSTKK-PHUQUOC", "id": "1702329333", "nm": "CÔNG TY TNHH KINGKONG PQ", "ad": "Tổ 8, Khu phố 10, Dương Đông, Đặc khu Phú Quốc, Tỉnh An Giang, Việt Nam", "gp": "ST", "nk": "", "nn": "KINGKONG - PHÚ QUỐC", "fc": "CT", "pv": "2026-10-03 00:00:00", "sh": "CH-ST-HC-XK"},
{"c": "CTXKSG-001", "id": "3703116109", "nm": "CÔNG TY TNHH VIỆT TIẾN FOODS", "ad": "Lô J42 đường DJ5 KĐT Mỹ Phước 3, Khu phố 3B, Phường Thới Hòa, Thành phố Hồ Chí Minh, Việt Nam", "gp": "XK", "nk": "", "nn": "", "fc": "CT", "pv": "2026-10-04 00:00:00", "sh": "CH-ST-HC-XK"},
{"c": "ĐNMNHN-0001", "id": "KHS0001", "nm": "Bán cho người tiêu dùng-Đặng Thị Thu Trang", "ad": "Phường Xuân Đỉnh, Thành Phố Hà Nội", "gp": "MN", "nk": "sỉ Đội Cấn", "nn": "THU TRANG ĐẶNG - HÀ NỘI (Sỉ Đội Cấn)", "fc": "ĐN", "pv": "HN", "sh": "MN"},
{"c": "ĐNMNHU-0001", "id": "046300009970", "nm": "Huỳnh Thị Mỹ Nhung (CCDC số: 046300009970)", "ad": "Thôn 3, xã Phú Vinh, thành phố Huế", "gp": "MN", "nk": "Nhung Huỳnh", "nn": "NHUNG HUỲNH - HUẾ", "fc": "ĐN", "pv": "HU", "sh": "MN"},
{"c": "CTMNSG-0001", "id": "KHS0002", "nm": "Bán cho người tiêu dùng-Nguyễn Thị Thanh", "ad": "402B Đường Lê Văn Lương, Phường Tân Hưng, Thành phố Hồ Chí Minh", "gp": "MN", "nk": "NGÂN BÍCH - PHÚ NHUẬN", "nn": "", "fc": "CT", "pv": "SG", "sh": "MN"},
{"c": "ĐNMNHY-0001", "id": "017301003150", "nm": "Nguyễn Tuyết Anh (CCCD số: 017301003150)", "ad": "Số 71 Minh Khai, Phường Thái Bình, tỉnh Hưng Yên", "gp": "MN", "nk": "Nguyễn Tuyết Anh", "nn": "NGUYỄN TUYẾT ANH - HƯNG YÊN (Sỉ Đội Cấn)", "fc": "ĐN", "pv": "HY", "sh": "MN"},
{"c": "ĐNMNHY-0002", "id": "034098015413", "nm": "HỘ KINH DOANH LINH ANH HAIR SALON", "ad": "Tại nhà ông Trần Thế Hệ, số nhà 71 Nguyễn Thị Minh Khai, tổ 8 Bồ Xuyên, Phường Thái Bình, Tỉnh Hưng Yên, Việt Nam", "gp": "MN", "nk": "Nguyễn Tuyết Anh 2", "nn": "", "fc": "ĐN", "pv": "HY", "sh": "MN"},
{"c": "ĐNMNHY-0003", "id": "", "nm": "Nguyễn Văn Linh (CCCD số: 034098015413)", "ad": "Số 71 Minh Khai, Phường Thái Bình, tỉnh Hưng Yên", "gp": "MN", "nk": "Nguyễn Tuyết Anh 3", "nn": "", "fc": "ĐN", "pv": "HY", "sh": "MN"},
{"c": "CTMNAG-0001", "id": "8416310831", "nm": "Nguyễn Hoàng Thiện", "ad": "xã Vĩnh Thạnh Trung, tỉnh An Giang", "gp": "MN", "nk": "TRIỆU MỸ HUYỀN - VỊNH TRE", "nn": "", "fc": "CT", "pv": "AG", "sh": "MN"},
{"c": "CTMNAG-0002", "id": "089093009004", "nm": "Huê Phước Thành (CCCD số: 089093009004)", "ad": "38c2 Thành Thái, phường Bình Đức, tỉnh An Giang", "gp": "MN", "nk": "THÀNH - LONG XUYÊN", "nn": "", "fc": "CT", "pv": "AG", "sh": "MN"},
{"c": "CTMNSG-0002", "id": "087301012836", "nm": "Huỳnh Ngọc Lan Như (CCCD số: 087301012836)", "ad": "Khu 3, Đặc khu Côn Đảo, Thành phố Hồ Chí Minh", "gp": "MN", "nk": "Huỳnh Như - Côn đảo", "nn": "", "fc": "CT", "pv": "SG", "sh": "MN"},
{"c": "CTMNSG-0003", "id": "079190014098", "nm": "Lê Hoàng Ngọc Thu (CCCD số: 079190014098)", "ad": "42/5/3 Thạnh Lộc 40, Khu phố 61, Phường An Phú Đông, Thành phố Hồ Chí Minh", "gp": "MN", "nk": "NGỌC THU - GÒ VẤP", "nn": "", "fc": "CT", "pv": "SG", "sh": "MN"},
{"c": "ĐNMNQI-0001", "id": "KHS0003", "nm": "Bán cho người tiêu dùng-Trương Tường Vy (TƯỜNG VY - QUẢNG NINH)", "ad": "Đường Hùng Vương, Phường Móng Cái 2, Tỉnh Quảng Ninh", "gp": "MN", "nk": "Tường Vy", "nn": "TƯỜNG VY - QUẢNG NINH", "fc": "ĐN", "pv": "QI", "sh": "MN"},
{"c": "ĐNMNPT-0001", "id": "KHS0004", "nm": "Bán cho người tiêu dùng-Nguyễn Thị Hồng Quyên", "ad": "Khu 1 (Chợ Tràng), Xã Phú Mỹ, Tỉnh Phú Thọ", "gp": "MN", "nk": "Sỉ Đội Cấn", "nn": "HỒNG QUYÊN - PHÚ THỌ (Sỉ Đội Cấn)", "fc": "ĐN", "pv": "PT", "sh": "MN"},
{"c": "ĐNMNHP-0001", "id": "", "nm": "Bán cho người tiêu dùng-Huỳnh Nguyễn Bảo Ngọc;", "ad": "Số 6, đường Bồ Đề Thủy, phường Thủy Nguyên, thành phố Hải Phòng", "gp": "MN", "nk": "Chè Huế Thủy Nguyên 1", "nn": "", "fc": "ĐN", "pv": "HP", "sh": "MN"},
{"c": "ĐNMNHP-0002", "id": "", "nm": "Bán hàng cho người tiêu dùng - Huỳnh Ngọc Quang", "ad": "Số 6, đường Bồ Đề Thủy, phường Thủy Nguyên, thành phố Hải Phòng", "gp": "MN", "nk": "Chè Huế Thủy Nguyên 2", "nn": "", "fc": "ĐN", "pv": "HP", "sh": "MN"},
{"c": "ĐNMNHP-0003", "id": "", "nm": "Bán hàng cho người tiêu dùng - Nguyễn Thị Mỹ Hạnh", "ad": "Số 6, đường Bồ Đề Thủy, phường Thủy Nguyên, thành phố Hải Phòng", "gp": "MN", "nk": "Chè Huế Thủy Nguyên 3", "nn": "", "fc": "ĐN", "pv": "HP", "sh": "MN"},
{"c": "ĐNMNHP-0004", "id": "", "nm": "Bán hàng cho người tiêu dùng - Nguyễn Tâm Thanh", "ad": "Số 6, đường Bồ Đề Thủy, phường Thủy Nguyên, thành phố Hải Phòng", "gp": "MN", "nk": "Chè Huế Thủy Nguyên 4", "nn": "", "fc": "ĐN", "pv": "HP", "sh": "MN"},
{"c": "ĐNMNHP-0005", "id": "KHS0005", "nm": "Bán cho người tiêu dùng-Nguyễn Thị Ngọc", "ad": "Ngã tư Thủy Sơn, phường Thủy Nguyên, thành phố Hải Phòng;", "gp": "MN", "nk": "Bảo Ngọc Hải Phòng", "nn": "", "fc": "ĐN", "pv": "HP", "sh": "MN"},
{"c": "CTMNSG-0004", "id": "079194022224", "nm": "HỘ KINH DOANH TUYỀN ĐẶNG (Đặng Thị Thanh Tuyền)", "ad": "78/35/10 Tôn Thất Thuyết, Khu Phố 6, Phường Xóm Chiếu, Thành phố Hồ Chí Minh, Việt Nam", "gp": "MN", "nk": "TUYỀN ĐẶNG - NGUYỄN BIỂU", "nn": "", "fc": "CT", "pv": "SG", "sh": "MN"},
{"c": "CTMNCT-0001", "id": "096184018412", "nm": "Nguyễn Thị Bích Trân (CCCD số: 096184018412)", "ad": "Gần Chợ số 6, Xã Vĩnh Trinh, Thành Phố Cần Thơ", "gp": "MN", "nk": "BÉ TRÂN", "nn": "BÉ TRÂN - CẦN THƠ", "fc": "CT", "pv": "CT", "sh": "MN"},
{"c": "ĐNMNHN-0002", "id": "KHS0006", "nm": "Bán cho người tiêu dùng-Trương Văn Quang", "ad": "Ngõ 211 đường Bát Khối, phường Long Biên, thành phố Hà Nội", "gp": "MN", "nk": "Sỉ Long Biên", "nn": "TIỆM CỦA EBE YEJUN - HN (Sỉ Long Biên)", "fc": "ĐN", "pv": "HN", "sh": "MN"},
{"c": "ĐNMNHP-0006", "id": "030097005967", "nm": "HỘ KINH DOANH THẢO MINIMART (NGUYỄN THỊ PHƯƠNG THẢO)", "ad": "Số nhà 55, đường Nguyễn Lương Bằng, Xã Thanh Miện, Thành phố Hải Phòng", "gp": "MN", "nk": "Sỉ Long Biên", "nn": "PHƯƠNG THẢO MINIMART - HẢI PHÒNG (Sỉ Long Biên)", "fc": "ĐN", "pv": "HP", "sh": "MN"},
{"c": "CTMNSG-0005", "id": "KHS0007", "nm": "Bán cho người tiêu dùng-Trần Thị Tuyết Nhung", "ad": "Đường Cao Đạt, phường Chợ Quán, thành phố Hồ Chí Minh", "gp": "MN", "nk": "Trần Thị Tuyết Nhung (trà sữa Bum - Nguyễn Biểu)", "nn": "", "fc": "CT", "pv": "SG", "sh": "MN"},
{"c": "ĐNMNQT-0001", "id": "044091012744", "nm": "Hoàng Huy Thành (CCCD số: 044091012744)", "ad": "Phường Đồng Hới, tỉnh Quảng Trị", "gp": "MN", "nk": "HOÀNG HUY THÀNH - CCCD SỐ: 044091012744 (Nga Ngắn)", "nn": "NGA NGẮN - QUẢNG TRỊ", "fc": "ĐN", "pv": "QT", "sh": "MN"},
{"c": "CTMNSG-0006", "id": "038195006858", "nm": "Trịnh Thị Hiền (CCCD số: 038195006858)", "ad": "358 Tô Ký , Phường Tân Chánh Hiệp, TP. Hồ Chí Minh", "gp": "MN", "nk": "TRỊNH THÚY HIỀN - GÒ VẤP", "nn": "", "fc": "CT", "pv": "SG", "sh": "MN"},
{"c": "CTMNCT-0002", "id": "094163007145", "nm": "Trương Thị Dung (CCCD số: 094163007145)", "ad": "211 Trần Hưng Đạo, Phường Vĩnh Châu, Thành phố Cần Thơ", "gp": "MN", "nk": "CHỊ CÚC - SÓC TRĂNG", "nn": "CHỊ CÚC - SÓC TRĂNG (CẦN THƠ)", "fc": "CT", "pv": "CT", "sh": "MN"},
{"c": "CTMNAG-0003", "id": "089096016882", "nm": "HỘ KINH DOANH HOÀNG LÂM", "ad": "Số 139, Đường Trần Hưng Đạo, Khóm 2, Xã Tri Tôn, Tỉnh An Giang, Việt Nam", "gp": "MN", "nk": "HOÀNG LÂM - TRI TÔN", "nn": "", "fc": "CT", "pv": "AG", "sh": "MN"},
{"c": "CTMNCT-0003", "id": "079179010111", "nm": "Đỗ Thị Ngọc Tiên (CCCD số: 079179010111)", "ad": "77 đường D4, Khu dân cư Hồng Loan, phường Cái Răng, thành phố Cần Thơ", "gp": "MN", "nk": "Thủy Tiên", "nn": "THỦY TIÊN - CẦN THƠ", "fc": "CT", "pv": "CT", "sh": "MN"},
{"c": "CTMNĐT-0001", "id": "082067020203", "nm": "NGUYỄN VĂN TÂM (CCCD số: 082067020203);", "ad": "Số 9/5a Lê Thị Hồng Gấm, Phường Thới Sơn, Tỉnh Đồng Tháp, Việt Nam", "gp": "MN", "nk": "TAM TONY - MỸ THO", "nn": "", "fc": "CT", "pv": "ĐT", "sh": "MN"},
{"c": "CTMNĐT-0002", "id": "", "nm": "Bán cho người tiêu dùng-Đinh Thị Thuỷ", "ad": "Số 9/5a Lê Thị Hồng Gấm, Phường Thới Sơn, Tỉnh Đồng Tháp, Việt Nam", "gp": "MN", "nk": "TAM TONY - MỸ THO 2", "nn": "", "fc": "CT", "pv": "ĐT", "sh": "MN"},
{"c": "CTMNSG-0007", "id": "KHS0008", "nm": "Bán cho người tiêu dùng-Nguyễn Thị Thanh;", "ad": "Số 135 Đội Cấn, phường Ngọc Hà, thành phố Hà Nội; 402B Lê Văn Lương, Phường Tân Hưng, Thành phố Hồ Chí Minh;", "gp": "MN", "nk": "THỦY MỘC - NGUYỄN BIỂU", "nn": "", "fc": "CT", "pv": "SG", "sh": "MN"},
{"c": "CTMNSG-0008", "id": "", "nm": "Nguyễn Thị Thanh Thuỷ (CCCD số: 056188009795);", "ad": "126 Hòn Khói, TDP Phú Thọ 2, Phường Đông Ninh Hòa, Khánh Hòa;", "gp": "MN", "nk": "THỦY MỘC - NGUYỄN BIỂU 2", "nn": "", "fc": "CT", "pv": "SG", "sh": "MN"},
{"c": "ĐNMNHU-0002", "id": "046093004031", "nm": "Nguyễn Doãn Phương Nam (CCCD số: 046093004031)", "ad": "Phường Phú Xuân, thành phố Huế", "gp": "MN", "nk": "Chút (Lào)", "nn": "CHÚT (LÀO) - HUẾ", "fc": "ĐN", "pv": "HU", "sh": "MN"},
{"c": "CTMNAG-0004", "id": "092195001263", "nm": "Dương Thị Diễm (CCCD số: 092195001263)", "ad": "Ấp Trường Phú B, xã Trường Long, phường Phong Điền, thành phố Cần Thơ, Việt Nam", "gp": "MN", "nk": "DIỄM PHÚ QUỐC", "nn": "", "fc": "CT", "pv": "AG", "sh": "MN"},
{"c": "ĐNMNPT-0002", "id": "025198010001", "nm": "Ngô Thị Thuỳ Dương (CCCD số: 025198010001);", "ad": "Số 161 Nguyễn Viết Xuân, phường Vĩnh Phúc, tỉnh Phú Thọ", "gp": "MN", "nk": "Sỉ Đội Cấn", "nn": "NGÔ THỊ THÙY DƯƠNG - PHÚ THỌ (Si Đội Cấn) 1", "fc": "ĐN", "pv": "PT", "sh": "MN"},
{"c": "ĐNMNPT-0003", "id": "", "nm": "Bán cho người tiêu dùng-Nguyễn Văn Khánh", "ad": "Số 161 Nguyễn Viết Xuân, phường Vĩnh Phúc, tỉnh Phú Thọ", "gp": "MN", "nk": "Sỉ Đội Cấn", "nn": "NGÔ THỊ THÙY DƯƠNG - PHÚ THỌ (Si Đội Cấn) 2", "fc": "ĐN", "pv": "PT", "sh": "MN"},
{"c": "CTMNSG-0009", "id": "KHS0009", "nm": "Bán cho người tiêu dùng-Trần Thị Thu Huyền", "ad": "Ấp Đông Lân, xã Bà Điểm, Thành phố Hồ Chí Minh", "gp": "MN", "nk": "Trần Thị thu Huyền (OMEGA)\n- Phú Nhuận", "nn": "", "fc": "CT", "pv": "SG", "sh": "MN"},
{"c": "ĐNMNPT-0004", "id": "036094019681", "nm": "Phạm Văn Đồng (CCCD số: 036094019681)", "ad": "Thôn Thanh Thiện, Xã Giao Minh, Tỉnh Ninh Bình, Việt Nam", "gp": "MN", "nk": "Phạm Văn Đồng - CCCD: 036094019681 (Đông Dola)", "nn": "ĐÔNG DOLA - HÀ NỘI (Phạm Văn Đồng)", "fc": "ĐN", "pv": "HN", "sh": "MN"},
{"c": "CTMNSG-0010", "id": "KHS0009", "nm": "Bán cho người tiêu dùng-Nguyễn Thị Thùy Dương", "ad": "Đường Cao Đạt, Phường Chợ Quán, Thành phố Hồ Chí Minh", "gp": "MN", "nk": "Nguyễn Thị Thùy Dương", "nn": "", "fc": "CT", "pv": "SG", "sh": "MN"},
{"c": "CTMNTN-0001", "id": "8462942105-001", "nm": "Nguyễn Thảo Vy", "ad": "Số 182, Khu phố 1, Võ Công Tồn, Xã Bến Lức, Tỉnh Tây Ninh, Việt Nam", "gp": "MN", "nk": "Nguyễn Thảo Vy (Thảo Vy)", "nn": "THẢO VY - LONG AN (TÂY NINH)", "fc": "CT", "pv": "TN", "sh": "MN"},
{"c": "ĐNMNQN-0001", "id": "062188000464", "nm": "Nguyễn Thị Thuỳ Dung (CCCD số: 062188000464 - THÙY DUNG - QUẢNG NGÃI)", "ad": "Số 164 Lê Hồng Phong, phường Kon Tum, tỉnh Quảng Ngãi", "gp": "MN", "nk": "Thùy Dung", "nn": "THÙY DUNG - QUẢNG NGÃI", "fc": "ĐN", "pv": "QN", "sh": "MN"},
{"c": "CTMNCM-0001", "id": "096087013455", "nm": "Ngô Kế Dư (CCCD số: 096087013455)", "ad": "Ấp 1, xã Trần Văn Thời, tỉnh Cà Mau", "gp": "MN", "nk": "KẾ DƯ - CÀ MAU", "nn": "", "fc": "CT", "pv": "CM", "sh": "MN"},
{"c": "ĐNMNĐN-0001", "id": "048194004103", "nm": "Đặng Hữu Triều Châu (CCCD số: 048194004103)", "ad": "127/6 Lý Tự Trọng, Phường Hải Châu, Thành phố Đà Nẵng", "gp": "MN", "nk": "Sỉ Xinh", "nn": "SỈ XINH - ĐÀ NẴNG", "fc": "ĐN", "pv": "ĐN", "sh": "MN"},
{"c": "ĐNMNĐN-0002", "id": "KHS0009", "nm": "Bán cho người tiêu dùng-Nguyễn Thị Thảo Nhi", "ad": "Thôn Kiệu Châu, xã Duy Xuyên, thành phố Đà Nẵng", "gp": "MN", "nk": "sỉ Maria Nhi", "nn": "MARIA NHI - ĐÀ NẴNG", "fc": "ĐN", "pv": "ĐN", "sh": "MN"},
{"c": "CTMNKH-0001", "id": "056193005382", "nm": "Nguyễn Thị Mỹ Vân (CCCD số: 056193005382)", "ad": "Số 112 Lạc Long Quân, phường Nha Trang, tỉnh Khánh Hoà", "gp": "MN", "nk": "VÂN - NHA TRANG", "nn": "", "fc": "CT", "pv": "KH", "sh": "MN"},
{"c": "CTMNSG-0011", "id": "8692341204-001", "nm": "HỘ KINH DOANH TIỆM HÀ LÊ (Lê Thị Thu Hà) CCCD: 064196008097", "ad": "Số nhà 165, đường Độc Lập, tổ 2, khu phố 1, Xã Phú Giáo, Thành phố Hồ Chí Minh, Việt Nam", "gp": "MN", "nk": "HÀ LÊ - NGUYỄN BIỂU", "nn": "", "fc": "CT", "pv": "SG", "sh": "MN"},
{"c": "ĐNMNHN-0004", "id": "037198000665", "nm": "Nguyễn Thị Mỹ Linh (CCCD số: 037198000665)", "ad": "297 Vũ Tông Phan, phường Khương Đình, thành phố Hà Nội", "gp": "MN", "nk": "Sỉ Hàng Bài (Phong Linh)", "nn": "PHONG LINH - HÀ NỘI (Sỉ Hàng Bài)", "fc": "ĐN", "pv": "HN", "sh": "MN"},
{"c": "CTMNCM-0002", "id": "096184018412", "nm": "Phạm Mỹ Thanh (CCCD số: 096184018412)", "ad": "Số 43 đường Dương Thị Cẩm Vân, ấp 4, xã Đầm Dơi, tỉnh Cà Mau", "gp": "MN", "nk": "DUY KHIÊM CÀ MAU", "nn": "", "fc": "CT", "pv": "CM", "sh": "MN"},
{"c": "CTMNSG-0012", "id": "KHS0009", "nm": "Bán cho người tiêu dùng-Nguyễn Thị Giang", "ad": "Phường Đông Hưng Thuận, thành phố Hồ Chí Minh", "gp": "MN", "nk": "DIỄM HƯƠNG - PHÚ NHUẬN", "nn": "", "fc": "CT", "pv": "SG", "sh": "MN"},
{"c": "ĐNMNHN-0005", "id": "001192020824", "nm": "Nguyễn Thị Hường (CCCD số: 001192020824)", "ad": "Ninh Hiệp, xã Phù Đổng, thành phố Hà Nội", "gp": "MN", "nk": "Sỉ Long Biên", "nn": "HƯƠNG CƯỜNG - HÀ NỘI (Sỉ Long Biên)", "fc": "ĐN", "pv": "HN", "sh": "MN"},
{"c": "CTMNSG-0013", "id": "KHS0009", "nm": "Bán cho người tiêu dùng-Nguyễn Thị Mỹ Linh", "ad": "320 Trần Bình Trọng, phường Chợ Quán, Thành phố Hồ Chí Minh", "gp": "MN", "nk": "MỸ LINH - NGUYỄN BIỂU", "nn": "", "fc": "CT", "pv": "SG", "sh": "MN"},
{"c": "ĐNMNĐN-0003", "id": "", "nm": "Bán cho người tiêu dùng-Cao Thị Thuý Hường", "ad": "Đường Ngô Quyền, Phường An Hải Nam, Thành phố Đà Nẵng", "gp": "MN", "nk": "Bảo Ngọc", "nn": "BẢO NGỌC - ĐÀ NẴNG", "fc": "ĐN", "pv": "ĐN", "sh": "MN"},
{"c": "CTMNSG-0014", "id": "KHS0009", "nm": "Bán cho người tiêu dùng-Nguyễn Thị Tâm", "ad": "Đường Âu Dương Lân, phường Chánh Hưng, Thành phố Hồ Chí Minh", "gp": "MN", "nk": "Pi House - NGUYỄN BIỂU", "nn": "", "fc": "CT", "pv": "SG", "sh": "MN"},
{"c": "CTMNLĐ-0001", "id": "060192007556", "nm": "Hồ Thị Hiếu (CCCD số: 060192007556 - Hiếu Hiếu - Nguyễn Biểu)", "ad": "178A Đường Tuyên Quang, Phường Phú Thủy, Tỉnh Lâm Đồng", "gp": "MN", "nk": "HIẾU HIẾU - NGUYỄN BIỂU", "nn": "HIẾU HIẾU - LÂM ĐỒNG (Nguyễn Biểu)", "fc": "CT", "pv": "LĐ", "sh": "MN"},
{"c": "CTMNSG-0015", "id": "KHS0009", "nm": "Bán cho người tiêu dùng-Lưu Văn Hùng", "ad": "337 Phan Đình Phùng, phường Cầu Kiệu, Thành Phố Hồ Chí Minh, Việt Nam", "gp": "MN", "nk": "MINH NGỌC - PHÚ NHUẬN", "nn": "", "fc": "CT", "pv": "SG", "sh": "MN"},
{"c": "CTMNĐO-0001", "id": "KHS0009", "nm": "Bán cho người tiêu dùng-Nguyễn Thị Tô Nguyên", "ad": "Phường Chơn Thành, tỉnh Đồng Nai", "gp": "MN", "nk": "NGUYÊN NGUYÊN - PHÚ NHUẬN", "nn": "NGUYÊN NGUYÊN - ĐỒNG NAI (Phú Nhuận)", "fc": "CT", "pv": "ĐO", "sh": "MN"},
{"c": "CTMNTN-0002", "id": "KHS0009", "nm": "Bán cho người tiêu dùng-Lê Ngọc Anh Tuấn", "ad": "phường Tân Ninh, tỉnh Tây Ninh", "gp": "MN", "nk": "Hậu Hậu - NB", "nn": "HẬU HẬU - TÂY NINH (Nguyễn Biểu)", "fc": "CT", "pv": "TN", "sh": "MN"},
{"c": "ĐNMNHP-0007", "id": "030191008729", "nm": "Nguyễn Thị Thu Huyền (CCCD số: 030191008729)", "ad": "146B Lê Thanh Nghị, Xã Gia Lộc, Thành phố Hải Phòng", "gp": "MN", "nk": "Huyền Hải Dương", "nn": "HUYỀN HẢI DƯƠNG - HẢI PHÒNG", "fc": "ĐN", "pv": "HP", "sh": "MN"},
{"c": "CTMNSG-0016", "id": "KHS0009", "nm": "Bán cho người tiêu dùng-Trần Thị Ngọc Anh", "ad": "Đường Cao Đạt, phường Chợ Quán, TP. Hồ Chí Minh", "gp": "MN", "nk": "PÉ BA - GÒ VẤP", "nn": "", "fc": "CT", "pv": "SG", "sh": "MN"},
{"c": "CTMNSG-0017", "id": "KHS0009", "nm": "Bán cho người tiêu dùng-Nguyễn Thị Thanh Hiền", "ad": "332 Quang Trung, phường Gò Vấp, TP. Hồ Chí Minh", "gp": "MN", "nk": "HIỀN NGUYỄN - PHÚ NHUẬN", "nn": "", "fc": "CT", "pv": "SG", "sh": "MN"},
{"c": "CTMNCM-0003", "id": "096199001113", "nm": "Dương Thị Vân Anh (CCCD số: 096199001113)", "ad": "Quán Dê Phố, Ấp 11, xã Sông Đốc, tỉnh Cà Mau.", "gp": "MN", "nk": "VÂN ANH - SÔNG ĐỐC", "nn": "", "fc": "CT", "pv": "CM", "sh": "MN"},
{"c": "CTMNSG-0018", "id": "KHS0009", "nm": "Bán cho người tiêu dùng-Nguyễn Thị Mỹ Trâm", "ad": "Khu phố Phước Trung, Xã Phước Hải, Thành phố Hồ Chí Minh", "gp": "MN", "nk": "NGUYỄN TRÂM - NGUYỄN BIỂU", "nn": "", "fc": "CT", "pv": "SG", "sh": "MN"},
{"c": "CTMNĐT-0003", "id": "1402207207", "nm": "CÔNG TY TNHH BIA BO BO (Tô Kim Bảo Trân)", "ad": "Số 4-6 Đường Phạm Hùng Dũng, khóm 2, Phường Hồng Ngự, Tỉnh Đồng Tháp, Việt Nam", "gp": "MN", "nk": "BẢO TRÂN - HỒNG NGỰ", "nn": "", "fc": "CT", "pv": "ĐT", "sh": "MN"},
{"c": "ĐNMNQT-0002", "id": "045197007564", "nm": "Lê Thị Cẩm Vân (CCCD số: 045197007564)", "ad": "Xã Cửa Tùng, tỉnh Quảng Trị", "gp": "MN", "nk": "Cẩm Vân", "nn": "CẨM VÂN - QUẢNG TRỊ", "fc": "ĐN", "pv": "QT", "sh": "MN"},
{"c": "ĐNMNĐN-0004", "id": "", "nm": "Bán cho người tiêu dùng-Trương Thị Kim Uyên", "ad": "Phường Cẩm Lệ, thành phố Đà Nẵng", "gp": "MN", "nk": "Kim Uyên", "nn": "KIM UYÊN - ĐÀ NẴNG", "fc": "ĐN", "pv": "ĐN", "sh": "MN"},
{"c": "CTMNCT-0004", "id": "", "nm": "Bán cho người tiêu dùng-Nguyễn Thanh Luân", "ad": "ấp Mỹ Quới, xã Hiệp Hưng, thành phố Cần Thơ", "gp": "MN", "nk": "THÀNH LUÂN - PHỤNG HIỆP", "nn": "THÀNH LUÂN - PHỤNG HIỆP (CẦN THƠ)", "fc": "CT", "pv": "CT", "sh": "MN"},
{"c": "CTMNVL-0001", "id": "083192002786", "nm": "Nguyễn Thị Thúy Ái (CCCD số: 083192002786)", "ad": "Thạnh B, Xã Đại Điền, Tỉnh Vĩnh Long, Việt Nam", "gp": "MN", "nk": "THÚY ÁI - BẾN TRE", "nn": "THÚY ÁI - BẾN TRE (VĨNH LONG)", "fc": "CT", "pv": "VL", "sh": "MN"},
{"c": "ĐNMNĐN-0005", "id": "", "nm": "Bán cho người tiêu dùng-Nguyễn Thị Kim Sanh", "ad": "Đường Phước Lý, phường Cẩm Lệ, thành phố Đà Nẵng", "gp": "MN", "nk": "Kim Sanh", "nn": "KIM SANH - ĐÀ NẴNG", "fc": "ĐN", "pv": "ĐN", "sh": "MN"},
{"c": "CTMNĐT-0004", "id": "087197007969", "nm": "Nguyễn Võ Khánh Linh (CCCD số: 087197007969)", "ad": "Đường Võ Thị Sáu, thị trấn Tràm Chim, huyện Tam Nông, tỉnh Đồng Tháp", "gp": "MN", "nk": "KHÁNH LINH TRÀM CHIM", "nn": "", "fc": "CT", "pv": "ĐT", "sh": "MN"},
{"c": "ĐNMNPT-0004", "id": "", "nm": "Bán cho người tiêu dùng-Nguyễn Thị Thu Hà", "ad": "Tổ 11, Khu 5, phường Thanh Miếu, tỉnh Phú Thọ", "gp": "MN", "nk": "Thu Hà", "nn": "THU HÀ - PHÚ THỌ", "fc": "ĐN", "pv": "PT", "sh": "MN"},
{"c": "CTMNCM-0004", "id": "095178007040", "nm": "Lê Hà Anh Tú (CCCD số: 095178007040)", "ad": "Ấp Long Thành, xã Phước Long, tỉnh Cà Mau", "gp": "MN", "nk": "HUỲNH LÊ PHÁT - PHƯỚC LONG", "nn": "", "fc": "CT", "pv": "CM", "sh": "MN"},
{"c": "CTMNSG-0019", "id": "077197002754", "nm": "Vũ Thị Trúc Trinh (CCCD số: 077197002754)", "ad": "Khu 2 đường Phan Chu Trinh, đặc khu Côn Đảo, Thành phố Hồ Chí Minh, Việt Nam", "gp": "MN", "nk": "TRINH CÔN ĐẢO - PN", "nn": "", "fc": "CT", "pv": "SG", "sh": "MN"},
{"c": "ĐNMNĐN-0006", "id": "", "nm": "Bán cho người tiêu dùng-Phạm Văn Khang", "ad": "705 Điện Biên Phủ, phường An Khê, thành phố Đà Nẵng", "gp": "MN", "nk": "PHẠM VĂN KHANG - PN", "nn": "PHẠM VĂN KHANG - ĐÀ NẴNG", "fc": "ĐN", "pv": "ĐN", "sh": "MN"},
{"c": "CTMNTN-0003", "id": "080165007510", "nm": "Trần Thị Ngọc Huệ", "ad": "07 ấp Chánh Hội, Xã Hậu Nghĩa, Tỉnh Tây Ninh", "gp": "MN", "nk": "SÁU HUỆ - NGUYỄN BIỂU", "nn": "SÁU HUỆ - TÂY NINH (Nguyễn Biểu)", "fc": "CT", "pv": "TN", "sh": "MN"},
{"c": "ĐNMNSL-0001", "id": "014300002515", "nm": "Lường Thùy Linh (CCCD số: 014300002515)", "ad": "Phường Tô Hiệu, tỉnh Sơn La", "gp": "MN", "nk": "Linh SuBee", "nn": "LINH SUBEE - SƠN LA", "fc": "ĐN", "pv": "SL", "sh": "MN"},
{"c": "ĐNMNHN-0006", "id": "001195005106", "nm": "Đinh Thị Mỹ Linh (CCCD số: 001195005106)", "ad": "Xóm 4, thôn Ninh Hiệp, xã Phù Đổng, Thành phố Hà Nội", "gp": "MN", "nk": "Sỉ Long Biên\nMỹ Linh- Long Biên", "nn": "MỸ LINH NINH HIỆP (Sỉ Long Biên)", "fc": "ĐN", "pv": "HN", "sh": "MN"},
{"c": "CTMNSG-0020", "id": "", "nm": "Bán cho người tiêu dùng-Hoàng Thị Oanh", "ad": "Số 9A, đường Tân Thới Nhất, phường Đông Hưng Thuận, Thành phố Hồ Chí Minh", "gp": "MN", "nk": "NHI NHI - PHÚ NHUẬN", "nn": "", "fc": "CT", "pv": "SG", "sh": "MN"},
{"c": "CTMNKH-0002", "id": "", "nm": "Bán cho người tiêu dùng-Đinh Lệ Thuỷ", "ad": "Ngã 3, xã Xuân Sơn, huyện Vạn Ninh, tỉnh Khánh Hòa", "gp": "MN", "nk": "ĐINH LỆ THỦY - KHÁNH HÒA", "nn": "", "fc": "CT", "pv": "KH", "sh": "MN"},
{"c": "CTMNSG-0021", "id": "091202000048", "nm": "HỘ KINH DOANH VYBAO BAKERY", "ad": "162/89 Xóm Chiếu, Phường Khánh Hội, Thành phố Hồ Chí Minh, Việt Nam", "gp": "MN", "nk": "NGUYỄN HOÀNG BẢO - NGUYỄN BIỂU", "nn": "", "fc": "CT", "pv": "SG", "sh": "MN"},
{"c": "ĐNMNHN-0007", "id": "", "nm": "Bán cho người tiêu dùng-Lương Thị Hằng", "ad": "Tòa nhà NO4, Khu đô thị Pháp Vân – Tứ Hiệp, phố Trần Thủ Độ, phường Yên Sở, thành phố Hà Nội", "gp": "MN", "nk": "Sỉ Hàng Bài", "nn": "HẰNG ASUKA - HÀ NỘI (Sỉ Hàng Bài)", "fc": "ĐN", "pv": "HN", "sh": "MN"},
{"c": "ĐNMNQN-0002", "id": "", "nm": "Bán cho người tiêu dùng-Lê Thị Bé Lũy", "ad": "Cây xăng Thanh Bình, tỉnh Quảng Ngãi", "gp": "MN", "nk": "Lê Thị Bủm", "nn": "LÊ THỊ BỦM - QUẢNG NGÃI", "fc": "ĐN", "pv": "QN", "sh": "MN"},
{"c": "ĐNMNHT-0001", "id": "042187014831", "nm": "Lê Thị Kim Yến (CCCD số: 042187014831)", "ad": "Số 04, Ngõ 16, Đường Đặng Dung, Phường Thành Sen, Tỉnh Hà Tĩnh.", "gp": "MN", "nk": "Lê Thị Kim Yến", "nn": "LÊ THỊ KIM YẾN - HÀ TĨNH", "fc": "ĐN", "pv": "HT", "sh": "MN"},
{"c": "ĐNMNCB-0001", "id": "004189009135", "nm": "Dương Thị Minh Huyền (CCCD số: 004189009135)", "ad": "Số nhà 65, phố Xuân Trường, phường Thục Phán, tỉnh Cao Bằng", "gp": "MN", "nk": "Sỉ Đội Cấn", "nn": "HUYỀN KIM THANH - CAO BẰNG (Sỉ Đội Cấn)", "fc": "ĐN", "pv": "CB", "sh": "MN"},
{"c": "CTMNSG-0022", "id": "038195006858", "nm": "Trịnh Thúy Hiền (CCCD số: 038195006858)", "ad": "358 Tô Ký, Phường Tân Chánh Hiệp, Thành phố Hồ Chí Minh", "gp": "MN", "nk": "TRỊNH THUÝ HIỀN - GÒ VẤP", "nn": "", "fc": "CT", "pv": "SG", "sh": "MN"},
{"c": "ĐNMNHN-0008", "id": "", "nm": "Bán cho người tiêu dùng-Nguyễn Thị Huyền Trang", "ad": "Phường Long Biên, Thành phố Hà Nội", "gp": "MN", "nk": "Sỉ Long Biên", "nn": "TRANG MILA - HÀ NỘI (Sỉ Long Biên)", "fc": "ĐN", "pv": "HN", "sh": "MN"},
{"c": "ĐNMNHN-0009", "id": "", "nm": "Bán cho người tiêu dùng-Nguyễn Thị Trang", "ad": "SN13/147/63, thôn Trung, xã Phù Đổng, thành phố Hà Nội", "gp": "MN", "nk": "Sỉ Đội Cấn", "nn": "MAI TRANG - HÀ NỘI (Sỉ Đội Cấn)", "fc": "ĐN", "pv": "HN", "sh": "MN"},
{"c": "CTMNSG-0023", "id": "", "nm": "Bán cho người tiêu dùng-Phạm Gia Tuệ", "ad": "Đường Nguyễn Biểu, Phường Chợ Quán, TP. Hồ Chí Minh", "gp": "MN", "nk": "BÁNH KEM THÙY DUNG - SỐ 10", "nn": "", "fc": "CT", "pv": "SG", "sh": "MN"},
{"c": "ĐNMNHN-0010", "id": "036197006013", "nm": "Phạm Thị Huệ (CCCD số: 036197006013)", "ad": "81h Ngõ 101 Đào Tấn, Phường Giảng Võ, Thành phố Hà Nội.", "gp": "MN", "nk": "Sỉ Đội Cấn", "nn": "NHƯ HUỆ - HÀ NỘI (Sỉ Đội Cấn)", "fc": "ĐN", "pv": "HN", "sh": "MN"},
{"c": "CTMNVL-0002", "id": "083190008911", "nm": "Hồ Thị Thuý Loan (CCCD số: 083190008911)", "ad": "596B16 Khu Phố 3, Nguyễn Thị Định, phường Phú Khương, tỉnh Vĩnh Long", "gp": "MN", "nk": "THÚY LOAN", "nn": "THÚY LONG - VĨNH LONG", "fc": "CT", "pv": "VL", "sh": "MN"},
{"c": "ĐNMNHN-0011", "id": "", "nm": "Bán cho người tiêu dùng-Nguyễn Thị Thu Hiền", "ad": "458 Minh Khai, Phường Vĩnh Tuy, Thành phố Hà Nội", "gp": "MN", "nk": "Sỉ Long Biên", "nn": "HIỀN HIỀN -HÀ NỘI (Sỉ Long Biên)", "fc": "ĐN", "pv": "HN", "sh": "MN"},
{"c": "ĐNMNPT-0005", "id": "025196009342", "nm": "Nguyễn Thị Nhung (CCCD số: 025196009342);", "ad": "Số nhà 215A, đường Tiên Dung, phường Thanh Miếu, tỉnh Phú Thọ.", "gp": "MN", "nk": "Mai Đoàn", "nn": "MAI ĐOÀN - PHÚ THỌ 1", "fc": "ĐN", "pv": "PT", "sh": "MN"},
{"c": "ĐNMNPT-0006", "id": "KHS0009", "nm": "Bán cho người tiêu dùng-Đoàn Thị Mai", "ad": "", "gp": "MN", "nk": "Mai Đoàn", "nn": "MAI ĐOÀN - PHÚ THỌ 2", "fc": "ĐN", "pv": "PT", "sh": "MN"},
{"c": "ĐNMNQI-0002", "id": "022196002549", "nm": "Nguyễn Thị Nguyệt Thanh (CCCD số: 022196002549)", "ad": "22 Trần Phú, phường Móng Cái 1, tỉnh Quảng Ninh", "gp": "MN", "nk": "NGUYỆT THANH MÓNG CÁI", "nn": "NGUYỆT THANH - QUẢNG NINH", "fc": "ĐN", "pv": "QI", "sh": "MN"},
{"c": "ĐNMNPT-0007", "id": "", "nm": "Bán cho người tiêu dùng-Trần Thị Trang", "ad": "Số 85 Vũ Trọng Phụng, phường Thanh Xuân Trung, thành phố Hà Nội", "gp": "MN", "nk": "Sỉ Đội Cấn", "nn": "SỈ ĐẮNG - PHÚ THỌ (Sỉ Đội Cấn)", "fc": "ĐN", "pv": "PT", "sh": "MN"},
{"c": "CTMNSG-0024", "id": "3701986139", "nm": "CÔNG TY TNHH PHONG THUẬN CƯỜNG", "ad": "Thửa đất số 1585, tờ bản đồ số 4, khu phố Khánh Lộc, Phường Tân Khánh, Thành phố Hồ Chí Minh, Việt Nam", "gp": "MN", "nk": "TÚ KIM LA - PHÚ NHUẬN", "nn": "", "fc": "CT", "pv": "SG", "sh": "MN"},
{"c": "CTMNAG-0005", "id": "091300002246", "nm": "Bùi Thị Kim Đang (CCCD số: 091300002246)", "ad": "Ấp Hòa Đông, Xã Đông Hòa, tỉnh An Giang", "gp": "MN", "nk": "KIM ĐANG - AN MINH", "nn": "", "fc": "CT", "pv": "AG", "sh": "MN"},
{"c": "ĐNMNTQ-0001", "id": "", "nm": "Bán cho người tiêu dùng-Vũ Thị Kim Thìn", "ad": "Thôn Vĩnh Lim, xã Chiêm Hóa, tỉnh Tuyên Quang", "gp": "MN", "nk": "Sỉ Hàng Bài", "nn": "KIM THÌN - TUYÊN QUANG (Sỉ Hàng Bài)", "fc": "ĐN", "pv": "TQ", "sh": "MN"},
{"c": "ĐNMNĐN-0007", "id": "035180006956", "nm": "HỘ KINH DOANH HƯƠNG VIỆT LOCAL FOOD", "ad": "Số 218 Nguyễn Chơn, Phường Hòa Khánh, Thành phố Đà Nẵng, Việt Nam", "gp": "MN", "nk": "Nguyễn Thị Hạnh", "nn": "NGUYỄN THỊ HẠNH - ĐÀ NẴNG", "fc": "ĐN", "pv": "ĐN", "sh": "MN"},
{"c": "CTMNSG-0025", "id": "", "nm": "Bán cho người tiêu dùng-Ngô Huỳnh Xuân Thảo", "ad": "Đường Thạnh Lộc 40, phường An Phú Đông, Tp.Hồ Chí Minh", "gp": "MN", "nk": "NGỌC THU - GÒ VẤP", "nn": "", "fc": "CT", "pv": "SG", "sh": "MN"},
{"c": "ĐNMNHT-0002", "id": "", "nm": "Bán cho người tiêu dùng-Nguyễn Thị Trinh", "ad": "Xóm 3, Xã Hà Linh, Tỉnh Hà Tĩnh", "gp": "MN", "nk": "Nguyễn Thị Trinh", "nn": "NGUYỄN THỊ TRINH - HÀ TĨNH", "fc": "ĐN", "pv": "HT", "sh": "MN"},
{"c": "CTMNSG-0026", "id": "", "nm": "Bán cho người tiêu dùng-Hà Tú Tân", "ad": "Phường Chợ Quán, Thành phố Hồ Chí Minh", "gp": "MN", "nk": "BAKERY HÂN HÂN - NGUYỄN BIỂU\nHÂN HÂN - NGUYỄN BIỂU", "nn": "", "fc": "CT", "pv": "SG", "sh": "MN"},
{"c": "CTMNKH-0003", "id": "056198009360", "nm": "Nguyễn Thị Kim Lan (CCCD số: 056198009360)", "ad": "Xã Diên Điền, tỉnh Khánh Hòa", "gp": "MN", "nk": "PÉ HAI - NGUYỄN BIỂU", "nn": "PÉ HAI - KHÁNH HÒA (Nguyễn Biểu)", "fc": "CT", "pv": "KH", "sh": "MN"},
{"c": "CTMNSG-0027", "id": "094196009966", "nm": "Thị Thanh Ni (CCCD số: 094196009966)", "ad": "Số 78 Đường Số 7, Phường Bình Hưng Hòa, Thành phố Hồ Chí Minh.", "gp": "MN", "nk": "Thị Thanh Ni ( Milk Thanh )- NGUYEN BIEU", "nn": "", "fc": "CT", "pv": "SG", "sh": "MN"},
{"c": "CTMNSG-0028", "id": "031097010339", "nm": "Nguyễn Văn Quang (CCCD số: 031097010339)", "ad": "Thôn Đại Hợp, xã Kiến Hải, thành phố Hải Phòng", "gp": "MN", "nk": "SỈ ANH - PHÚ NHUẬN", "nn": "", "fc": "CT", "pv": "SG", "sh": "MN"},
{"c": "ĐNMNHN-0012", "id": "026198004866", "nm": "Nguyễn Thị Ánh Tuyết (CCCD số: 026198004866)", "ad": "Phường Đội Cấn, Thành Phố Hà Nội", "gp": "MN", "nk": "Mun Nguyễn - 69", "nn": "MUN NGUYỄN (Sỉ Đội Cấn)", "fc": "ĐN", "pv": "HN", "sh": "MN"},
{"c": "CTMNSG-0029", "id": "0319163094", "nm": "CÔNG TY TNHH LY LY BAKERY (LY LY -NGUYỄN BIỂU)", "ad": "Số 575/1A Hoàng Sa, Phường Xuân Hòa, Thành phố Hồ Chí Minh, Việt Nam", "gp": "MN", "nk": "LY LY -NGUYỄN BIỂU", "nn": "", "fc": "CT", "pv": "SG", "sh": "MN"},
{"c": "ĐNMNHN-0013", "id": "", "nm": "Bán cho người tiêu dùng-Nguyễn Thị Nở", "ad": "Số 17, ngõ 3 phố Trịnh Tố Tâm, phường Bồ Đề, thành phố Hà Nội", "gp": "MN", "nk": "Sỉ Long Biên", "nn": "HỒ THỊ BÍCH VÂN - HÀ NỘI (Sỉ Long Biên)", "fc": "ĐN", "pv": "HN", "sh": "MN"},
{"c": "CTMNSG-0030", "id": "", "nm": "Bán cho người tiêu dùng-Đỗ Thị Hồng Thu", "ad": "Phường Chợ Quán, Thành phố Hồ Chí Minh", "gp": "MN", "nk": "HỒNG THU YOGA - NGUYỄN BIỂU", "nn": "", "fc": "CT", "pv": "SG", "sh": "MN"},
{"c": "CTMNCM-0005", "id": "096091007619", "nm": "Lâm Thái Bảo (CCCD số: 096091007619)", "ad": "47 An Dương Vương, khóm 2 Năm Căn, Tỉnh Cà Mau, Việt Nam", "gp": "MN", "nk": "HUY THÁI - CÀ MAU", "nn": "", "fc": "CT", "pv": "CM", "sh": "MN"},
{"c": "CTMNLĐ-0002", "id": "", "nm": "Bán cho người tiêu dùng-Tạ Thị Thúy Xuân", "ad": "Ngã 3 Gia Viễn, Cát Tiên, Đạ Huoai, Lâm Đồng", "gp": "MN", "nk": "THÚY XUÂN - PHÚ NHUẬN", "nn": "THÚY XUÂN - LÂM ĐỒNG (Phú Nhuận)", "fc": "CT", "pv": "LĐ", "sh": "MN"},
{"c": "CTMNTN-0004", "id": "072192004243", "nm": "Hoàng Thị Ngọc Giàu (CCCD số: 072192004243)", "ad": "Ấp 1, xã Tân Châu, tỉnh Tây Ninh", "gp": "MN", "nk": "BONNY - NGUYỄN BIỂU", "nn": "BONNY - TÂY NINH (Nguyễn Biểu)", "fc": "CT", "pv": "TN", "sh": "MN"},
{"c": "CTMNVL-0003", "id": "", "nm": "Bán cho người tiêu dùng-Phan Thị Bích Tuyền", "ad": "Xã An Định, tỉnh Vĩnh Long", "gp": "MN", "nk": "Phan Bích Tuyền - mỏ Cày Nam", "nn": "BÍCH TUYỀN - VĨNH LONG (MỎ CÀY NAM)", "fc": "CT", "pv": "VL", "sh": "MN"},
{"c": "ĐNMNĐN-0008", "id": "", "nm": "Bán cho người tiêu dùng-Nguyễn Thị Mai Phương", "ad": "Phường Cẩm Lệ, Thành Phố Đà Nẵng", "gp": "MN", "nk": "Phương Mai", "nn": "PHƯƠNG MAI - ĐÀ NẴNG", "fc": "ĐN", "pv": "ĐN", "sh": "MN"},
{"c": "ĐNMNHN-0014", "id": "", "nm": "Bán cho người tiêu dùng-Nguyễn Thu Hà", "ad": "Ngõ 91, Nguyễn Văn Trỗi, Phường Phương Liệt, Thành phố Hà Nội", "gp": "MN", "nk": "Sỉ Hàng Bài", "nn": "HOÀNG TUẤN - HÀ NỘI (Sỉ Hàng Bài)", "fc": "ĐN", "pv": "HN", "sh": "MN"},
{"c": "ĐNMNNA-0001", "id": "", "nm": "Bán cho người tiêu dùng-HUANG CHIN LUAN", "ad": "Xã Tân Kỳ, tỉnh Nghệ An", "gp": "MN", "nk": "Thanh Huyền", "nn": "THANH HUYỀN - NGHỆ AN", "fc": "ĐN", "pv": "NA", "sh": "MN"},
{"c": "ĐNMNTQ-0002", "id": "008192002042", "nm": "Nguyễn Thị Trang (CCCD số: 008192002042)", "ad": "Xóm 5, Nhà máy Z129, xã Nhữ Khê, tỉnh Tuyên Quang", "gp": "MN", "nk": "Sỉ Đội Cấn", "nn": "MAI TRANG - TUYÊN QUANG (Sỉ Đội Cấn)", "fc": "ĐN", "pv": "TQ", "sh": "MN"},
{"c": "CTMNĐL-0001", "id": "054195008577", "nm": "Trần Trương Kiều Linh (CCCD số: 054195008577)", "ad": "147 Nguyễn Thái Học, phường Tuy Hòa, tỉnh Đắk Lắk", "gp": "MN", "nk": "THÀNH HOÀI - NGUYỄN BIỂU", "nn": "THANH HOÀI - ĐẮK LẮK (Nguyễn Biểu)", "fc": "CT", "pv": "ĐL", "sh": "MN"},
{"c": "CTMNVL-0004", "id": "084192007384", "nm": "HỘ KINH DOANH BÁNH KEM KHÁNH LY - ĐỊA ĐIỂM KINH DOANH TIỆM BÁNH ĐÔ LA (Huỳnh Khánh Phụng Tiên)", "ad": "26 Nguyễn Đáng, khóm 3, phường Long Đức,, Tỉnh Vĩnh Long, Việt Nam", "gp": "MN", "nk": "BÁNH KEM KHÁNH LY - TRÀ VINH", "nn": "BÁNH KEM KHÁNH LY - TRÀ VINH (VĨNH LONG)", "fc": "CT", "pv": "VL", "sh": "MN"},
{"c": "ĐNMNSL-0002", "id": "", "nm": "Bán hàng cho người tiêu dùng - Nguyễn Thị Ánh Tuyết", "ad": "Tổ 2, phường Chiềng Sinh, Thành Phố Sơn La, Tỉnh Sơn La", "gp": "MN", "nk": "Sỉ Đội Cấn", "nn": "ÁNH TUYẾT - SƠN LA (Sỉ Đội Cấn)", "fc": "ĐN", "pv": "SL", "sh": "MN"},
{"c": "ĐNMNHP-0008", "id": "030191008729", "nm": "Nguyễn Thị Thu Huyền (CCCD số: 030191008729)", "ad": "146B Lê Thanh Nghị, xã Gia Lộc, thành phố Hải Phòng", "gp": "MN", "nk": "Huyền Hải Dương", "nn": "HUYỀN HẢI DƯƠNG", "fc": "ĐN", "pv": "HP", "sh": "MN"},
{"c": "ĐNMNHP-0009", "id": "", "nm": "Bán cho người tiêu dùng-Hà Thị Trang", "ad": "Cổng KCN LG – Tràng Duệ, phường An Phong, thành phố Hải Phòng", "gp": "MN", "nk": "Hà Trang", "nn": "HÀ TRANG - HẢI PHÒNG", "fc": "ĐN", "pv": "HP", "sh": "MN"},
{"c": "ĐNMNSL-0003", "id": "", "nm": "Bán cho người tiêu dùng-Nguyễn Thị Thủy", "ad": "Phường Mộc Châu, tỉnh Sơn La", "gp": "MN", "nk": "Sỉ Hàng Bài", "nn": "BÍCH THỦY - SƠN LA (Sỉ Hàng Bài)", "fc": "ĐN", "pv": "SL", "sh": "MN"},
{"c": "CTMNCT-0005", "id": "092307000645", "nm": "Phạm Thị Thúy Ngân (CCCD số: 092307000645)", "ad": "Xã Thới Hưng, Thành phố Cần Thơ", "gp": "MN", "nk": "THÚY NGÂN - CỜ ĐỎ", "nn": "THÚY NGÂN - CỜ ĐỎ (CẦN THƠ)", "fc": "CT", "pv": "CT", "sh": "MN"},
{"c": "ĐNMNQT-0003", "id": "3101150991", "nm": "CÔNG TY TNHH MTV NHƯ VY;\nTrần Tiến Dũng", "ad": "Số 89 Trần Hưng Đạo, Phường Đồng Hới, Tỉnh Quảng Trị, Việt Nam; 08 Trần Quang Khải, Phường Đồng Hới, tỉnh Quảng Trị", "gp": "MN", "nk": "Cẩm Nguyễn", "nn": "CẨM NGUYỄN - QUẢNG TRỊ", "fc": "ĐN", "pv": "QT", "sh": "MN"},
{"c": "CTMNCT-0006", "id": "093177010121", "nm": "Nguyễn Thị Kim Giác (CCCD số: 093177010121)", "ad": "Khu vực 8, Phường Ngã Bảy, Thành phố Cần Thơ", "gp": "MN", "nk": "Nguyễn Thị Kim Giác - CCCD số: 093177010121(Trạm dừng chân Hiệp Lợi 3)", "nn": "Nguyễn Thị Kim Giác -(Trạm dừng chân Hiệp Lợi 3) (CẦN THƠ)", "fc": "CT", "pv": "CT", "sh": "MN"},
{"c": "CTMNSG-0031", "id": "", "nm": "Bán cho người tiêu dùng-Mai Thị Linh", "ad": "Đường Phan Đình Phùng, phường Đức Nhuận, Thành phố Hồ Chí Minh", "gp": "MN", "nk": "MAI LINH - PHÚ NHUẬN", "nn": "", "fc": "CT", "pv": "SG", "sh": "MN"},
{"c": "ĐNMNHN-0015", "id": "008199006141", "nm": "Ngô Thuỳ Linh (CCCD số: 008199006141)", "ad": "Phố Hàng Bài, Phường Cửa Nam, Thành phố Hà Nội", "gp": "MN", "nk": "Sỉ Hàng Bài", "nn": "SỈ LINH - HÀ NỘI (Sỉ Hàng Bài)", "fc": "ĐN", "pv": "HN", "sh": "MN"},
{"c": "CTMNKH-0004", "id": "056194014348", "nm": "HỘ KINH DOANH BÁNH KEM TRÂN TRÂN (sỉ Huyền Vân Nha Trang)", "ad": "829 đường 23/10, thôn Phú Thạnh, Phường Tây Nha Trang, Tỉnh Khánh Hòa, Việt Nam", "gp": "MN", "nk": "Huyền Vân Nha Trang", "nn": "", "fc": "CT", "pv": "KH", "sh": "MN"},
{"c": "CTMNCM-0006", "id": "0948853202", "nm": "Bán cho người tiêu dùng-Nguyễn Anh Thư (Sỉ Thư Em)", "ad": "Xã Gành Hào, Tỉnh Cà Mau", "gp": "MN", "nk": "THƯ EM - BẠC LIÊU", "nn": "", "fc": "CT", "pv": "CM", "sh": "MN"},
{"c": "ĐNMNTA-0001", "id": "0335908999", "nm": "Bán cho người tiêu dùng - Hứa Thị Hồng Nhung (Hồng Nhung - Thái Nguyên)", "ad": "Thôn 5, xã Chợ Đồn, tỉnh Thái Nguyên", "gp": "MN", "nk": "", "nn": "HỒNG NHUNG - THÁI NGUYÊN", "fc": "ĐN", "pv": "TA", "sh": "MN"},
{"c": "ĐNMNHP-0010", "id": "0348055999", "nm": "Bán cho người tiêu dùng - Nguyễn Khánh Linh 0348055999", "ad": "Quán Gỏi xã Kẻ Sặt tp Hải Phòng", "gp": "MN", "nk": "Tùng 0348055999", "nn": "", "fc": "ĐN", "pv": "HP", "sh": "MN"},
{"c": "CTMNSG-0032", "id": "079168009225", "nm": "Dương Thị Thanh Hằng (CCCD số: 079168009225)", "ad": "500/42 Đoàn Văn Bơ, Phường Khánh Hội, Thành phố Hồ Chí Minh", "gp": "MN", "nk": "HẢO PHỤNG", "nn": "", "fc": "CT", "pv": "SG", "sh": "MN"},
{"c": "ĐNMNLC-0001", "id": "", "nm": "Bán cho người tiêu dùng-Nguyễn Lan Phương 0848521333", "ad": "Thị Trấn Mù Cang Chải, Xã Mù Cang Chải, tỉnh Lào Cai", "gp": "MN", "nk": "Nguyễn Minh Tú", "nn": "", "fc": "ĐN", "pv": "LC", "sh": "MN"},
{"c": "CTMNSG-0033", "id": "", "nm": "Bán cho người tiêu dùng - Lê Hà Thị Ngọc Mai (điện thoại 0903959720)", "ad": "12.07 Lô A, Chung Cư Bàu Cát 2, Phường Bảy Hiền, Thành phố Hồ Chí Minh", "gp": "MN", "nk": "Sỉ Mai - Nguyễn Biểu", "nn": "", "fc": "CT", "pv": "SG", "sh": "MN"},
{"c": "CTMNSG-0034", "id": "086198009352", "nm": "Dương Ngọc Nhung (CCCD: 086198009352-CHỊ THU HÀ HCM)", "ad": "L77 Đinh Tiên Hoàng, Phường Tân Hạnh, Tỉnh Vĩnh Long", "gp": "CH", "nk": "Dương Ngọc Nhung (CHỊ THU HÀ HCM)", "nn": "", "fc": "CT", "pv": "SG", "sh": "MN"},
{"c": "ĐNMNQT-0004", "id": "0357035973", "nm": "Bán cho người tiêu dùng - Phan Thị Ngọc Ánh", "ad": "Thôn Nội Hoa, Xã Đông Trạch, Tỉnh Quảng Trị", "gp": "MN", "nk": "Sỉ Ngọc Ánh", "nn": "NGỌC ÁNH - QUẢNG TRỊ", "fc": "ĐN", "pv": "QT", "sh": "MN"},
{"c": "ĐNMNSL-0004", "id": "", "nm": "Bán cho người tiêu dùng - Nguyễn Trang Nhung", "ad": "Tổ 2, phường Chiềng Sinh, Thành Phố Sơn La, Tỉnh Sơn La", "gp": "MN", "nk": "", "nn": "ÁNH TUYẾT - SƠN LA (Sỉ Đội Cấn)", "fc": "ĐN", "pv": "SL", "sh": "MN"},
{"c": "CTMNSG-0035", "id": "", "nm": "Bán cho người tiêu dùng - Mai Trúc Phương", "ad": "", "gp": "MN", "nk": "Cô Út - Bay Malay", "nn": "", "fc": "CT", "pv": "SG", "sh": "MN"},
{"c": "CTMNSG-0036", "id": "", "nm": "Bán cho người tiêu dùng - Chen Ssu-Yu", "ad": "Phường Đông Hưng Thuận, Thành phố Hồ Chí Minh", "gp": "MN", "nk": "SỈ DIỄM HƯƠNG", "nn": "", "fc": "CT", "pv": "SG", "sh": "MN"},
{"c": "CTDL1SG-0001", "id": "086198009352", "nm": "Dương Ngọc Nhung (CCCD: 086198009352-CHỊ THU HÀ HCM)", "ad": "L77 Đinh Tiên Hoàng, Phường Tân Hạnh, Tỉnh Vĩnh Long", "gp": "DL1", "nk": "Dương Ngọc Nhung (CHỊ THU HÀ HCM)", "nn": "", "fc": "CT", "pv": "SG", "sh": "DL1"},
{"c": "CTDL1SG-0002", "id": "086193001513", "nm": "HỘ KINH DOANH BÁNH SẦU RIÊNG THU HÀ", "ad": "250/3 Phan Anh, Phường Phú Thạnh, Thành phố Hồ Chí Minh, Việt Nam", "gp": "DL1", "nk": "HỘ KINH DOANH BÁNH SẦU RIÊNG THU HÀ (CHỊ HÀ HCM)", "nn": "", "fc": "CT", "pv": "SG", "sh": "DL1"},
{"c": "ĐNDL1NA-0001", "id": "040172025048", "nm": "Cao Thị Huệ (CCCD số: 040172025048-Cửa hàng Nghệ An)", "ad": "Thôn 7, xã Diễn Kỳ, huyện Diễn Châu, tỉnh Nghệ An", "gp": "DL1", "nk": "", "nn": "", "fc": "ĐN", "pv": "NA", "sh": "DL1"}
];

const DSM_STOP = ['ban','cho','nguoi','tieu','dung','si','cong','ty','tnhh','mtv','ho','kinh','doanh','hkd','cua','hang',
  'bao','nhieu','dia','chi','khach','sieu','thi','chinh','the','cp','co','phan','tm','dv','khu','pho','phuong','xa','tinh','thanh','tp','quan','huyen','so','duong','to','ap','khom'];

function _dsmTok_(s) {
  return _stripDiacriticsServer(s || '').replace(/\(.*?\)/g, ' ').replace(/[^a-z0-9]+/g, ' ').trim().split(/\s+/)
    .filter(function (x) { return x.length > 1 && DSM_STOP.indexOf(x) === -1; });
}
function _dsmSim_(a, b) {
  var A = _dsmTok_(a), B = _dsmTok_(b);
  if (!A.length || !B.length) return 0;
  var inter = 0; A.forEach(function (x) { if (B.indexOf(x) !== -1) inter++; });
  return inter / Math.max(Math.min(A.length, B.length), Math.ceil(Math.max(A.length, B.length) * 0.5));
}
function _dsmExact_(a, b) {
  var x = _stripDiacriticsServer(a || '').replace(/[^a-z0-9]+/g, ' ').trim();
  var y = _stripDiacriticsServer(b || '').replace(/[^a-z0-9]+/g, ' ').trim();
  return !!x && x === y;
}
function _dsmDig_(v) { return (v == null ? '' : v).toString().replace(/\D/g, ''); }
function _dsmIdEq_(a, b) {
  a = _dsmDig_(a); b = _dsmDig_(b);
  if (!a || !b) return false;
  return a === b || (a.length >= 10 && b.length >= 10 && a.substring(0, 10) === b.substring(0, 10));
}
function _dsmCleanName_(s) {
  s = (s == null ? '' : s).toString().trim();
  s = s.replace(/^bán\s+(hàng\s+)?cho\s+người\s+tiêu\s+dùng\s*[-–:]?\s*/i, '');
  return s.replace(/\s*\((?:CCCD|CCDC|CMND|điện thoại)[^)]*\)\s*/ig, ' ').replace(/\s+/g, ' ').trim();
}
function _dsmParseCode_(code) {
  var m = _normCode_(code).match(/^(CT|ĐN)(HCTM|HCGC|DL1|MN|CH|ST|XK|HC)/);
  if (!m) return null;
  var group = m[2], prov = '';
  if (group === 'MN' || group === 'CH' || group === 'DL1') prov = _normCode_(code).substr(m[0].length, 2);
  return { factory: m[1], group: group, prov: prov };
}
// Tách cột "Mã KH/MST/CCCD" của file: phone | oldCode | tax | cccd
function _dsmClassify_(idRaw, texts) {
  var out = { phone: '', tax: '', cccd: '', oldCode: '' };
  var s = (idRaw || '').toString().replace(/\s+/g, '');
  var blob = texts.join(' ');
  var pm = blob.match(/(?:^|[^\d])(0\d{9,10})(?!\d)/);
  if (pm) out.phone = pm[1];
  if (!s) return out;
  if (/^KHS\d+$/i.test(s)) { out.oldCode = s.toUpperCase(); return out; }
  if (/^[A-Za-zĐđ]/.test(s)) return out;
  if (s.indexOf('-') !== -1) { out.tax = s; return out; }
  var d = s.replace(/\D/g, '');
  if (out.phone && d === out.phone) { out.phone = d; return out; }   // cột này thực chất là SĐT
  if (d.length === 11) d = '0' + d;
  if (d.length === 12 || d.length === 9) out.cccd = d; else if (d) out.tax = d;
  return out;
}

function doiSoatMa_XemTruoc() { return _doiSoatMa_(false); }
function doiSoatMa_ApDung() { return _doiSoatMa_(true); }

function _doiSoatMa_(apply) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sh = _dataKhachSiSheet(), ss = sh.getParent(), W = DKS_TOTAL_COLS;
    var last = sh.getLastRow();
    var rows = last > 1 ? sh.getRange(2, 1, last - 1, W).getValues() : [];
    var C = DKS_COL;

    // --- chỉ mục dữ liệu hiện có ---
    var info = rows.map(function (r) {
      return {
        code: _normCode_(r[C.CODE - 1]),
        names: [r[C.NAME - 1], r[C.INVOICE_NAME - 1], r[C.COMPANY - 1], r[C.DESCRIPTION - 1]].filter(Boolean).map(String),
        phone: _dsmDig_(r[C.PHONE - 1]), tax: _dsmDig_(r[C.TAX - 1]), cccd: _dsmDig_(r[C.CCCD - 1]),
        addr: (r[C.ADDRESS - 1] || '').toString(), old: _normCode_(r[C.OLD_CODE - 1])
      };
    });
    var byCode = {};
    info.forEach(function (x, i) { if (x.code) { (byCode[x.code] = byCode[x.code] || []).push(i); } });

    // mã cũ KHS xuất hiện đúng 1 lần trong file mới được dùng để khớp
    var oldCount = {};
    DSM_FILE.forEach(function (k) { var o = /^KHS\d+$/i.test(k.id || '') ? k.id.toUpperCase() : ''; if (o) oldCount[o] = (oldCount[o] || 0) + 1; });

    var report = [], plan = [], claimed = {}, usedCode = {};
    var cnt = { ok: 0, fix: 0, assign: 0, add: 0, skip: 0 };
    var addRows = [], codeMap = {};   // codeMap: mã cũ trong hệ thống -> mã mới

    DSM_FILE.forEach(function (k) {
      var fcode = _normCode_(k.c);
      var disp = k.nn || k.nk || _dsmCleanName_(k.nm);
      var texts = [k.nm, k.nk, k.nn].filter(Boolean);
      var cl = _dsmClassify_(k.id, texts);
      var fileNames = texts.concat([disp]);
      var parsed = _dsmParseCode_(fcode);
      var warn = '';
      if (!parsed) warn = 'Mã không đúng cấu trúc [Xưởng][Nhóm]...';
      else if (k.gp && k.gp !== parsed.group && !(k.gp === 'HC' && /^HC/.test(parsed.group)) && !(parsed.group === 'DL1'))
        warn = 'Nhóm trong file (' + k.gp + ') khác nhóm trong mã (' + parsed.group + ') - kiểm tra lại file';

      if (usedCode[fcode] != null) {
        report.push([fcode, disp, 'DUP_FILE', '', '', '', '', 'Mã bị trùng ở nhiều dòng trong file (' + k.sh + ') - sửa file rồi chạy lại. ' + warn]);
        cnt.skip++; return;
      }

      // ---- tìm ứng viên ----
      var cands = {};
      var add = function (i, score, how) { if (!cands[i] || cands[i].score < score) cands[i] = { score: score, how: how }; };
      var nameSim = function (x) { var s = 0; x.names.forEach(function (rn) { fileNames.forEach(function (fn) { s = Math.max(s, _dsmSim_(fn, rn)); }); }); return s; };
      var nameExact = function (x) { return x.names.some(function (rn) { return fileNames.some(function (fn) { return _dsmExact_(fn, rn); }); }); };
      var holders = byCode[fcode] || [], holderNote = '';
      holders.forEach(function (i) {
        var x = info[i], s = nameSim(x), as = _dsmSim_(k.ad, x.addr);
        var ok = (cl.phone && x.phone === _dsmDig_(cl.phone)) || (cl.cccd && _dsmIdEq_(cl.cccd, x.cccd)) || (cl.tax && _dsmIdEq_(cl.tax, x.tax)) || s >= 0.5 || as >= 0.6;
        if (ok) add(i, 100, 'đang giữ mã này'); else holderNote = 'Mã đang do dòng ' + (i + 2) + ' (' + (rows[i][C.NAME - 1] || '') + ') giữ - khác người';
      });
      for (var i = 0; i < info.length; i++) {
        var x = info[i];
        if (cl.phone && x.phone && x.phone === _dsmDig_(cl.phone)) add(i, 90, 'SĐT');
        var idHit = (cl.cccd && (_dsmIdEq_(cl.cccd, x.cccd) || _dsmIdEq_(cl.cccd, x.tax))) || (cl.tax && (_dsmIdEq_(cl.tax, x.tax) || _dsmIdEq_(cl.tax, x.cccd)));
        var s = nameSim(x);
        if (idHit && s >= 0.5) add(i, 80, 'CCCD/MST + tên');
        else if (idHit && !x.names.length) add(i, 55, 'CCCD/MST');
        var as = _dsmSim_(k.ad, x.addr);
        if (nameExact(x) && as >= 0.5) add(i, 70, 'tên + địa chỉ');
        else if (s >= 0.8 && as >= 0.5) add(i, 60, 'tên giống + địa chỉ giống');
        if (cl.oldCode && oldCount[cl.oldCode] === 1 && x.old === cl.oldCode && s >= 0.3) add(i, 50, 'mã cũ ' + cl.oldCode);
      }
      var keys = Object.keys(cands).map(Number).sort(function (a, b) { return cands[b].score - cands[a].score; });
      var best = keys.length ? keys[0] : -1;
      if (keys.length > 1 && cands[keys[0]].score === cands[keys[1]].score) {
        report.push([fcode, disp, 'AMBIGUOUS', keys.map(function (q) { return q + 2; }).join(', '), '', '', '', 'Nhiều dòng cùng độ tin cậy - xử lý tay. ' + warn]);
        cnt.skip++; return;
      }
      if (best >= 0 && claimed[best] && claimed[best] !== fcode) {
        report.push([fcode, disp, 'CONFLICT', best + 2, info[best].code, rows[best][C.NAME - 1], cands[best].how, 'Dòng này đã khớp với mã ' + claimed[best] + ' trong file - 1 khách không thể có 2 mã. ' + warn]);
        cnt.skip++; return;
      }
      if (best < 0) {
        if (holders.length) {
          report.push([fcode, disp, 'CONFLICT', holders.map(function (q) { return q + 2; }).join(', '), fcode, rows[holders[0]][C.NAME - 1], '', holderNote + ' - không tự thêm mới vì sẽ trùng mã. ' + warn]);
          cnt.skip++; return;
        }
        usedCode[fcode] = -2;
        var arr = new Array(W).fill('');
        arr[C.NAME - 1] = disp; arr[C.ADDRESS - 1] = k.ad; arr[C.INVOICE_NAME - 1] = _dsmCleanName_(k.nm) === disp ? '' : k.nm;
        if (cl.phone) arr[C.PHONE - 1] = cl.phone; if (cl.tax) arr[C.TAX - 1] = cl.tax; if (cl.cccd) arr[C.CCCD - 1] = cl.cccd;
        if (cl.oldCode) arr[C.OLD_CODE - 1] = cl.oldCode;
        if (k.nk && !_dsmExact_(k.nk, disp)) arr[C.DESCRIPTION - 1] = k.nk;
        arr[C.CODE - 1] = fcode;
        arr[C.FACTORY - 1] = (parsed && parsed.factory) || k.fc; arr[C.PROV - 1] = (parsed && parsed.prov) || k.pv; arr[C.GROUP - 1] = (parsed && parsed.group) || k.gp;
        addRows.push(arr); cnt.add++;
        report.push([fcode, disp, 'ADD', '', '', '', '', 'Chưa có khách này trong hệ thống - thêm mới. ' + warn]);
        return;
      }
      // ---- khớp được 1 dòng ----
      claimed[best] = fcode; usedCode[fcode] = best;
      var cur = info[best].code, r = rows[best];
      var status;
      if (cur === fcode) { status = 'OK'; cnt.ok++; }
      else if (!cur) { status = 'ASSIGN'; cnt.assign++; }
      else { status = 'FIX'; cnt.fix++; codeMap[cur] = fcode; }
      if (holders.length && holders.indexOf(best) === -1) {
        // mã file đang do người khác giữ -> không thể gán
        report.push([fcode, disp, 'CONFLICT', best + 2, cur, r[C.NAME - 1], cands[best].how, holderNote + ' - cần xử lý tay (đổi mã người kia trước). ' + warn]);
        claimed[best] = null; usedCode[fcode] = null; if (status === 'OK') cnt.ok--; else if (status === 'ASSIGN') cnt.assign--; else { cnt.fix--; delete codeMap[cur]; }
        cnt.skip++; return;
      }
      if (status !== 'OK') {
        r[C.CODE - 1] = fcode;
        if (parsed) { r[C.FACTORY - 1] = parsed.factory; r[C.PROV - 1] = parsed.prov; r[C.GROUP - 1] = parsed.group; }
      } else if (parsed) {
        if (!r[C.FACTORY - 1]) r[C.FACTORY - 1] = parsed.factory;
        if (!r[C.PROV - 1] && parsed.prov) r[C.PROV - 1] = parsed.prov;
        if (!r[C.GROUP - 1]) r[C.GROUP - 1] = parsed.group;
      }
      var fill = function (col, val) { if (val && !(r[col - 1] || '').toString().trim()) r[col - 1] = val; };
      fill(C.ADDRESS, k.ad); fill(C.PHONE, cl.phone); fill(C.TAX, cl.tax); fill(C.CCCD, cl.cccd); fill(C.OLD_CODE, cl.oldCode);
      fill(C.INVOICE_NAME, k.nm);
      report.push([fcode, disp, status, best + 2, cur, r[C.NAME - 1], cands[best].how, status === 'OK' ? ('Đã đúng mã. ' + warn) : ('Mã hệ thống ' + (cur || '(trống)') + ' -> ' + fcode + '. ' + warn)]);
    });

    // --- mã trong hệ thống mà file không có: chỉ liệt kê để biết ---
    var fileCodes = {}; DSM_FILE.forEach(function (k) { fileCodes[_normCode_(k.c)] = true; });
    var extra = 0;
    info.forEach(function (x, i) {
      if (x.code && !fileCodes[x.code] && !claimed[i] && _dsmParseCode_(x.code)) {
        extra++;
        report.push([x.code, rows[i][C.NAME - 1], 'NOT_IN_FILE', i + 2, x.code, rows[i][C.NAME - 1], '', 'Mã có trong hệ thống nhưng KHÔNG có trong file New Code (có thể do cart tự cấp) - đối chiếu tay']);
      }
    });

    // --- ghi báo cáo ---
    var rp = ss.getSheetByName('KetQuaDoiSoatMa');
    if (!rp) rp = ss.insertSheet('KetQuaDoiSoatMa'); else rp.clear();
    var head = ['Mã chuẩn (file)', 'Tên', 'Kết quả', 'Dòng DataKhachSi', 'Mã hiện tại', 'Tên hiện tại', 'Cách khớp', 'Ghi chú'];
    rp.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#fde68a');
    if (report.length) { rp.getRange(2, 1, report.length, 1).setNumberFormat('@'); rp.getRange(2, 5, report.length, 1).setNumberFormat('@'); rp.getRange(2, 1, report.length, head.length).setValues(report); }
    rp.setFrozenRows(1);

    var backup = '', tracked = 0;
    if (apply) {
      backup = 'DataKhachSi_BACKUP_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
      sh.copyTo(ss).setName(backup);
      var fmt = [C.PHONE, C.TAX, C.CCCD, C.CODE, C.OLD_CODE];
      if (rows.length) {
        fmt.forEach(function (c) { sh.getRange(2, c, rows.length, 1).setNumberFormat('@'); });
        sh.getRange(2, 1, rows.length, W).setValues(rows);
      }
      if (addRows.length) {
        var st = sh.getLastRow() + 1;
        fmt.forEach(function (c) { sh.getRange(st, c, addRows.length, 1).setNumberFormat('@'); });
        sh.getRange(st, 1, addRows.length, W).setValues(addRows);
      }
      // đổi mã cũ -> mới trong theo dõi hóa đơn (cột Mã KH = cột 5)
      try {
        var inv = ss.getSheetByName('TheoDoiHoaDon');
        if (inv && inv.getLastRow() > 1 && Object.keys(codeMap).length) {
          var rng = inv.getRange(2, 5, inv.getLastRow() - 1, 1), vals = rng.getValues();
          vals.forEach(function (v) { var c = _normCode_(v[0]); if (c && codeMap[c]) { v[0] = codeMap[c]; tracked++; } });
          rng.setNumberFormat('@'); rng.setValues(vals);
        }
      } catch (e2) { Logger.log('Không cập nhật được TheoDoiHoaDon: ' + e2); }
    }
    var msg = (apply ? 'ĐÃ ÁP DỤNG' : 'XEM TRƯỚC (chưa sửa data)') + ': đúng sẵn ' + cnt.ok + ' | gán mã cho khách chưa có mã ' + cnt.assign + ' | SỬA mã sai ' + cnt.fix +
      ' | thêm mới ' + cnt.add + ' | bỏ qua/xung đột ' + cnt.skip + ' | mã hệ thống không có trong file ' + extra +
      (apply ? (' | hóa đơn đã đổi mã: ' + tracked + ' | sao lưu: ' + backup) : '') + '. Chi tiết: sheet KetQuaDoiSoatMa.';
    Logger.log(msg);
    try { SpreadsheetApp.getActive().toast(msg, 'Đối soát mã khách', 20); } catch (e3) {}
    return msg;
  } finally { lock.releaseLock(); }
}