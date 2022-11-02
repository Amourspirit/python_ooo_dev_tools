from __future__ import annotations
import sys
from pathlib import Path
from draw_gradient import DrawGradient, GradientKind
from ooodev.utils.file_io import FileIO
import argparse


def args_add(parser: argparse.ArgumentParser) -> None:
    # usage for default start.py -k
    parser.add_argument(
        "-k",
        "--kind",
        const="fill",
        nargs="?",
        dest="kind",
        choices=[e.value for e in GradientKind],
        help="Kind of gradient to display (default: %(default)s)",
    )
    parser.add_argument(
        "-w",
        "--width",
        dest="width",
        type=int,
        required=False,
        help="Optional - Width of gradient to display",
    )
    parser.add_argument(
        "-t",
        "--height",
        dest="height",
        type=int,
        required=False,
        help="Optional - Height of gradient to display",
    )
    parser.add_argument(
        "-x",
        "--x-pos",
        dest="x",
        type=int,
        required=False,
        help="Optional - X position of gradient to display",
    )
    parser.add_argument(
        "-y",
        "--y-pos",
        dest="y",
        type=int,
        required=False,
        help="Optional - Y position of gradient to display",
    )
    parser.add_argument(
        "-s",
        "--start-color",
        dest="start_color",
        required=False,
        help="Optional - Start color of gradient to display",
    )
    parser.add_argument(
        "-e",
        "--end-color",
        dest="end_color",
        required=False,
        help="Optional - End color of gradient to display",
    )
    parser.add_argument(
        "-a",
        "--angle",
        dest="angle",
        type=int,
        required=False,
        help="Optional - angle of gradient to display",
    )
    parser.add_argument(
        "-n",
        "--gradient-name",
        dest="gradient_name",
        required=False,
        help="Optional - Name of gradient to display",
    )


# region main()
def main() -> int:
    if len(sys.argv) == 1:
        sys.argv.append("-k")
    # create parser to read terminal input
    parser = argparse.ArgumentParser(description="main")

    # add args to parser
    args_add(parser=parser)

    # read the current command line args
    args = parser.parse_args()

    fnm = Path("tests/fixtures/image/crazy_blue.jpg")
    p = FileIO.get_absolute_path(fnm)
    if not p.exists():
        fnm = Path("../../../../tests/fixtures/image/crazy_blue.jpg")
        p = FileIO.get_absolute_path(fnm)
    if not p.exists():
        raise FileNotFoundError("Unable to find path to crazy_blue.jpg")

    kind = GradientKind(args.kind)

    cv = DrawGradient(gradient_kind=kind, gradient_fnm=fnm)
    if args.width:
        cv.width = args.width
    if args.height:
        cv.height = args.height
    if args.x:
        cv.x = args.x
    if args.y:
        cv.y = args.y
    if args.start_color:
        try:
            cv.start_color = int(args.start_color)
        except ValueError:
            cv.start_color = int(args.start_color, 16)
    if args.end_color:
        try:
            cv.end_color = int(args.end_color)
        except ValueError:
            cv.end_color = int(args.end_color, 16)
    if args.angle:
        cv.angle = args.angle
    if args.gradient_name:
        cv.name_gradient = args.gradient_name
    cv.main()
    return 0


# endregion main()

if __name__ == "__main__":
    SystemExit(main())
