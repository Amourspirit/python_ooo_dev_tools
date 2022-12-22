#!/usr/bin/env python
# coding: utf-8
import argparse
import sys
import os
from pathlib import Path
from src.cmds import uno_lnk, run_auto, manage_env_cfg
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


def _args_cmd_toggle_evn(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "-t",
        "--toggle-env",
        help="Toggle virtual environment to and from LibreOffice environment",
        action="store_true",
        dest="toggle_env",
        default=False,
    )
    parser.add_argument(
        "-u",
        "--uno-env",
        help="Displayes if the current Virtual Environment is UNO Environment.",
        action="store_true",
        dest="uno_env",
        default=False,
    )
    parser.add_argument(
        "-c",
        "--custom-env",
        help="Set a custom environment. cfg file must must be manually configured.",
        action="store",
        dest="cusom_env",
        required=False,
    )


def _args_action_cmd_link(a_parser: argparse.ArgumentParser, args: argparse.Namespace) -> None:
    if not (args.add or args.remove):
        a_parser.error("No action requested, add --add or --remove")
    if args.add:
        uno_lnk.add_links(args.src_dir)
    elif args.remove:
        uno_lnk.remove_links()


def _args_action_cmd_auto(a_parser: argparse.ArgumentParser, args: argparse.Namespace) -> None:
    pargs = str(args.process_file).split()
    if sys.platform == "win32":
        run_auto.run_lo_py(*pargs)
    else:
        run_auto.run_py(*pargs)


def _args_action_cmd_toggle_env(a_parser: argparse.ArgumentParser, args: argparse.Namespace) -> None:
    if args.uno_env:
        if manage_env_cfg.is_env_uno_python():
            print("UNO Environment")
        else:
            print("NOT a UNO Environment")
        return
    if args.toggle_env:
        manage_env_cfg.toggle_cfg()
        return
    if args.cusom_env:
        manage_env_cfg.toggle_cfg(suffix=args.cusom_env)


def _args_process_cmd(a_parser: argparse.ArgumentParser, args: argparse.Namespace) -> None:
    if args.command == "cmd-link":
        _args_action_cmd_link(a_parser=a_parser, args=args)
    elif args.command == "auto":
        _args_action_cmd_auto(a_parser=a_parser, args=args)
    elif args.command == "env":
        _args_action_cmd_toggle_env(a_parser=a_parser, args=args)
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

    if sys.platform == "win32":
        cmd_env_toggle = subparser.add_parser(
            name="env",
            help="Manage Virtual Environment configuration.",
        )
        _args_cmd_toggle_evn(parser=cmd_env_toggle)

    cmd_auto = subparser.add_parser(name="auto", help="Run an automation script")

    _args_cmd_auto(parser=cmd_auto)

    # region Read Args
    args = parser.parse_args()
    # endregion Read Args
    _args_process_cmd(a_parser=parser, args=args)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
