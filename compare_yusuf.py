# Сравнение результата Юсуфа с эталоном «готовый»
#
#   python compare_yusuf.py samples/out_yusuf.xlsx "samples/Готовый Юсуф №1643.xlsx"
#
import sys

import convert_nakladnaya as core


def main():
    if len(sys.argv) < 3:
        print("Usage: python compare_yusuf.py <результат.xlsx> <эталон_готовый.xlsx>")
        sys.exit(2)

    diffs = core.sravnit_yusuf(sys.argv[1], sys.argv[2])
    if diffs:
        print("Расхождений: %s" % len(diffs))
        for line in diffs[:40]:
            print(line)
        sys.exit(1)
    print("OK: 0 расхождений")


if __name__ == "__main__":
    main()
