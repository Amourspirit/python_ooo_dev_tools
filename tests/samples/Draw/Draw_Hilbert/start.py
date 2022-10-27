from draw_hilbert import DrawHilbert
import sys

# region main()
def main() -> int:
    level = 4
    if len(sys.argv) == 2:
        level = int(sys.argv[1])
    dh = DrawHilbert(level)
    dh.draw()
    return 0


# endregion main()

if __name__ == "__main__":
    SystemExit(main())
