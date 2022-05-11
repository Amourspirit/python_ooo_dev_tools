#!/usr/bin/env python
# coding: utf-8
import argparse
import sys
import os
from pathlib import Path
from src.cmds import uno_lnk, run_auto
from src.build.build import Builder, BuilderArgs
from src.utils import util

# region parser
# region        Create Parsers


def _create_parser(name: str) -> argparse.ArgumentParser:
    return argparse.ArgumentParser(description=name)


# endregion     Create Parsers

# region        process arg command
def _args_cmd_link(parser: argparse.ArgumentParser) -> None:
    add_grp = parser.add_argument_group()
    add_grp.add_argument(
        "-a",
        "--add",
        help="Add uno links to virtual environment.",
        action="store_true",
        dest="add",
        default=False,
    )

    add_grp.add_argument(
        "-s",
        "--uno-src",
        help="Optional source directory that contains uno.py and unohelper.py. If ommited then defaults are used.",
        action="store",
        dest="src_dir",
        default=None,
    )
    parser.add_argument(
        "-r",
        "--remove",
        help="Remove uno links to virtual environment.",
        action="store_true",
        dest="remove",
        default=False,
    )


def _args_cmd_auto(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "-p",
        "--process",
        help="Path to file to run such as ex/auto/writer/hello_world/main.py",
        action="store",
        dest="process_file",
        required=True,
    )


def _args_cmd_build(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "-c",
        "--config",
        help="Json config file that contains build info.",
        action="store",
        dest="config_json",
        required=True,
    )
    parser.add_argument(
        "-v",
        "--verbose",
        help="If True some print statements will be made in the terminal.",
        action="store_true",
        dest="verbose",
        default=False,
    )
    embed_grp = parser.add_argument_group()
    embed_grp.add_argument(
        "-e",
        "--embed",
        help="If True script embeded script will be embed in a doc.",
        action="store_true",
        dest="embed",
        default=False,
    )
    embed_grp.add_argument(
        "-s",
        "--embed-src",
        help="Source document to embed script into. If omitted then default to interanl odt file.",
        action="store",
        dest="embed_src",
        default=None,
    )


def _args_action_cmd_link(
    a_parser: argparse.ArgumentParser, args: argparse.Namespace
) -> None:
    if not (args.add or args.remove):
        a_parser.error("No action requested, add --add or --remove")
    if args.add:
        uno_lnk.add_links(args.src_dir)
    elif args.remove:
        uno_lnk.remove_links()


def _args_action_cmd_build(
    a_parser: argparse.ArgumentParser, args: argparse.Namespace
) -> None:
    bargs = BuilderArgs(
        config_json=args.config_json,
        embed_in_doc=bool(args.embed),
        embed_doc=args.embed_src,
        allow_print=bool(args.verbose),
    )
    builder = Builder(args=bargs)
    _valid = builder.build()
    if _valid == False:
        print("Build Failed")
    else:
        print("Build Success")


def _args_action_cmd_auto(
    a_parser: argparse.ArgumentParser, args: argparse.Namespace
) -> None:
    pargs = str(args.process_file).split()
    if sys.platform == "win32":
        run_auto.run_lo_py(*pargs)
    else:
        run_auto.run_py(*pargs)


def _args_process_cmd(
    a_parser: argparse.ArgumentParser, args: argparse.Namespace
) -> None:
    if args.command == "cmd-link":
        _args_action_cmd_link(a_parser=a_parser, args=args)
    elif args.command == "build":
        _args_action_cmd_build(a_parser=a_parser, args=args)
    elif args.command == "auto":
        _args_action_cmd_auto(a_parser=a_parser, args=args)
    else:
        a_parser.print_help()


# endregion        process arg command
# endregion parser


def _main() -> int:
    # for debugging
    args = "build -e --config ex/general/apso_console/config.json --embed-src ex/general/apso_console/apso_example.odt"
    # args = "auto -p ex/auto/writer/hello_world/main.py"
    sys.argv.extend(args.split())
    return main()


def main() -> int:
    os.environ["project_root"] = str(Path(__file__).parent)
    os.environ["env-site-packages"] = str(util.get_site_packeges_dir())
    parser = _create_parser("main")
    subparser = parser.add_subparsers(dest="command")

    if os.name != "nt":
        # linking is not useful in Windows.
        cmd_link = subparser.add_parser(
            name="cmd-link",
            help="Add/Remove links in virtual environments to uno files.",
        )
        _args_cmd_link(parser=cmd_link)

    cmd_build = subparser.add_parser(name="build", help="Build a script")

    cmd_auto = subparser.add_parser(name="auto", help="Run an automation script")

    _args_cmd_build(parser=cmd_build)
    _args_cmd_auto(parser=cmd_auto)

    # region Read Args
    args = parser.parse_args()
    # endregion Read Args
    _args_process_cmd(a_parser=parser, args=args)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
