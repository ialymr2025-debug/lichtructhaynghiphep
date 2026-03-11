export const DEFAULT_STAFF = [
  ["Trưởng ca",        "Nguyễn Tiến Danh",  "Tạ Văn Hà",          "Nguyễn Văn Trường", "Bùi Chí Thanh",      "Hoàng Tấn Hùng"],
  ["Trực TTĐK",        "Vũ Đức Cường",       "Lê Trí Dũng",        "Nguyễn Văn Toàn",   "Mai Xuân Sơn",       "Trần Trung Liêm"],
  ["Trực chính điện",  "Phan Văn Hùng",      "Ngô Xuân Đoàn",      "Nguyễn Yên Nam",    "Nguyễn Trung Chính", "Lê Hướng"],
  ["Trực phụ điện",    "Hoàng Ngọc Ân",      "Nguyễn Ngọc Hùng",   "Nguyễn Phỉ Được",   "Trần Lê Hưng",       "Trương Đình Thắng"],
  ["Trực chính máy",   "Nguyễn Tấn Phước",   "Lê Văn Dân",         "Phạm Văn Toàn",     "Phan Ngọc Hùng",     "Nguyễn Khắc Học"],
  ["Trực phụ máy",     "Phạm Văn Von",       "Thái Trần Hoàng Vũ", "Lê Văn Tích",       "Puih Thăn",          "Kiều Cao Khởi"],
  ["TC TBA 500 kV",    "Võ Quang Minh",      "Trần Hữu Thuận",     "Lê Thành Cao",      "Bùi Ngọc Thuận",     "Phạm Hồng Thắng"],
  ["Trực OPY",         "Ngô Xuân Vỹ",        "Trần Văn Thiên",     "Hoàng Văn Thăng",   "Trần Tiễn Quân",     "A Ran"],
  ["Trực CNN",         "Phạm Văn Mạnh",      "Hà Văn Chăn",        "Lê Thế Đàm",        "Cù Minh Trung",      "Nguyễn Vinh Quang"],
  ["Trưởng kíp",       "Nguyễn Lâm Tiến",    "Đỗ Văn Anh",         "Trịnh Xuân An",     "Tào Trọng Thi",      "Vũ Huy Hùng"],
  ["Trực chính GM",    "Nguyễn Hồng Quang",  "Trần Nhật Huy",      "Võ Thành Trung",    "Phùng Ngọc Tú",      "Nguyễn Thành Nguyên"],
  ["Trực phụ điện MR", "Nguyễn Thanh Phong", "Lê Hoài Bảo",        "Lê Vũ Minh Trung",  "Phạm Đình Đức",      "Lê Trọng Toàn"],
  ["Trực phụ cơ MR",   "Nguyễn Quang Minh",  "Rmah Thắng",         "Phạm Thanh Tùng",   "Nguyễn Văn Trung",   "Huỳnh Viết Vương"]
];

export const SHIFTS = [
  ["C","O","K","O","N","O"],
  ["O","C","O","K","O","N"],
  ["N","O","C","O","K","O"],
  ["O","N","O","C","O","K"],
  ["K","O","N","O","C","O"]
];

export const BASE_DATE = new Date(2025, 9, 3);

export const RULES: Record<number, Record<string, { k: number }>> = {
  1: { N: { k: 4 }, C: { k: 2 }, K: { k: 3 } },
  2: { N: { k: 5 }, C: { k: 3 }, K: { k: 4 } },
  3: { N: { k: 1 }, C: { k: 4 }, K: { k: 5 } },
  4: { N: { k: 2 }, C: { k: 5 }, K: { k: 1 } },
  5: { N: { k: 3 }, C: { k: 1 }, K: { k: 2 } }
};
