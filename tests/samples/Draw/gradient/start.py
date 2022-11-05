from __future__ import annotations
import sys
from pathlib import Path
from ooodev.office.draw import DrawingHatchingKind, DrawingBitmapKind, DrawingGradientKind
from draw_gradient import DrawGradient, GradientKind
from ooodev.utils.file_io import FileIO
from ooodev.utils.color import CommonColor
import argparse


def args_add(parser: argparse.ArgumentParser) -> None:
    # usage for default start.py -k
    parser.add_argument(
        "-k",
        "--kind",
        const="hatch",
        nargs="?",
        dest="kind",
        choices=[e.value for e in GradientKind],
        help="Kind of gradient to display (default: %(default)s)",
    )
    parser.add_argument(
        # by add name as lower and setting type to type=str.lower
        # the user can input upper or lower and it will work fine.
        # note that type=str.lower has no brackets.
        "--hatch-kind",
        default="green_30_degrees",
        nargs="?",
        dest="hatch_kind",
        type=str.lower,
        choices=[e.name.lower() for e in DrawingHatchingKind],
        help="Kind of hatch gradient to display (default: %(default)s)",
    )
    parser.add_argument(
        # by add name as lower and setting type to type=str.lower
        # the user can input upper or lower and it will work fine.
        # note that type=str.lower has no brackets.
        "--bitmap-kind",
        default="floral",
        nargs="?",
        dest="bitmap_kind",
        type=str.lower,
        choices=[e.name.lower() for e in DrawingBitmapKind],
        help="Kind of bitmap gradient to display (default: %(default)s)",
    )
    parser.add_argument(
        "--gradient-kind",
        default="neon_light",
        nargs="?",
        dest="gradient_kind",
        type=str.lower,
        choices=[e.name.lower() for e in DrawingGradientKind],
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
        cv.start_color = CommonColor.from_str(args.start_color)
    if args.end_color:
        cv.end_color = CommonColor.from_str(args.end_color)
    if args.angle:
        cv.angle = args.angle
    if args.gradient_kind:
        cv.name_gradient = DrawingGradientKind.from_str(args.gradient_kind)
    if args.hatch_kind:
        cv.hatch_gradient = DrawingHatchingKind.from_str(args.hatch_kind)
    if args.bitmap_kind:
        cv.bitmap_gradient = DrawingBitmapKind.from_str(args.bitmap_kind)
    cv.main()
    return 0


# endregion main()

if __name__ == "__main__":
    SystemExit(main())
