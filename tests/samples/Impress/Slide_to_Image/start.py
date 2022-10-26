from __future__ import annotations
import argparse
import sys

from slide_2_image import Slide2Image


def args_add(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "-f",
        "--file",
        help="File path of input file",
        action="store",
        dest="file_path",
        required=True,
    )
    parser.add_argument(
        "-i",
        "--idx",
        help="Optonal index of slide to convert to image. Default: %(default)i",
        action="store",
        dest="idx",
        type=int,
        default=0,
    )
    parser.add_argument(
        "-o",
        "--out_fmt",
        help="Extension of the converted file. Default: %(default)s",
        action="store",
        dest="output_format",
        default="jpeg",
    )
    parser.add_argument(
        "-d",
        "--output_dir",
        help="Optional output Directory. Defaults to temporary dir sub folder.",
        action="store",
        dest="out_dir",
        default="",
    )


def main() -> int:
    # create parser to read terminal input
    parser = argparse.ArgumentParser(description="main")

    # add args to parser
    args_add(parser=parser)

    # read the current command line args
    if len(sys.argv) == 1:
        parser.print_help()
        return 0
    args = parser.parse_args()
    sl = Slide2Image(fnm=args.file_path, idx=args.idx, img_fmt=args.output_format, out_dir=args.out_dir)
    sl.main()

    return 0


if __name__ == "__main__":
    SystemExit(main())
