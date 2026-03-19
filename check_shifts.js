
const BASE_DATE = new Date(2025, 9, 1);
const SHIFTS = [
  ["N","C","O","K","O"],
  ["K","O","N","C","O"],
  ["C","O","K","O","N"],
  ["O","N","C","O","K"],
  ["O","K","O","N","C"]
];

function xacDinhCa(ngay, kip) {
  const diff = Math.floor((ngay.getTime() - BASE_DATE.getTime()) / 86400000);
  const cycleLen = SHIFTS[0].length;
  return SHIFTS[kip - 1][((diff % cycleLen) + cycleLen) % cycleLen];
}

const d1 = new Date(2026, 2, 9);
console.log("2026-03-09:");
for (let k = 1; k <= 5; k++) {
  console.log(`Kip ${k}: ${xacDinhCa(d1, k)}`);
}

const d2 = new Date(2026, 2, 14);
console.log("\n2026-03-14:");
for (let k = 1; k <= 5; k++) {
  console.log(`Kip ${k}: ${xacDinhCa(d2, k)}`);
}
