#!/usr/bin/env python3
"""Convert a Beamer presentation to PowerPoint. See --help for the options."""

import sys

from beamer2pptx.cli import main

if __name__ == "__main__":
    sys.exit(main())
