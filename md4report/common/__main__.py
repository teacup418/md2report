import argparse
import os
from .handlers.add_info_header import worker
from .utils import get_metadata

parser = argparse.ArgumentParser()

parser.add_argument(
    "src",
    type=str,
    help="source file path"
)

parser.add_argument(
    "config",
    type=str,
    help="config's name"
)

args = parser.parse_args()

metadata = get_metadata(args.src)
target = os.path.splitext(args.src)[0] + ".docx"

worker(metadata, target)
