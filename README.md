# klword

# klword

[![PyPI](https://img.shields.io/pypi/v/klword)](https://pypi.org/project/klword/)
[![Python](https://img.shields.io/pypi/pyversions/klword)](https://pypi.org/project/klword/)
[![License](https://img.shields.io/github/license/poggiotzx/klword)](LICENSE)

A lightweight **Word report generation library** built on top of
**docxtpl** and **python-docx**.

`klword` provides a simple API for generating structured Word reports
from templates.\
It is designed for automated workflows such as testing reports,
simulation reports, and batch document generation.

------------------------------------------------------------------------

# Features

-   Template-based Word report generation
-   Simple and structured API
-   Insert text, images, and tables easily
-   Rich text styling support
-   Container-based document assembly
-   Compatible with standard `.docx` templates

------------------------------------------------------------------------

# Installation

Install from PyPI:

``` bash
pip install klword
```

Or install from source:

``` bash
git clone https://github.com/yourname/klword.git
cd klword
pip install .
```

------------------------------------------------------------------------

# Quick Start

``` python
from klword import WordAPI
from klword.styles import BODY_STYLE, make_rich_text

api = WordAPI("template.docx")

text = make_rich_text(
    "This text is generated automatically.",
    BODY_STYLE
)

container = api.new_container()
container.add_text(text)

api.render({
    "content": container
})

api.save("report.docx")
```

------------------------------------------------------------------------

# Template Example

Create a Word template (`template.docx`) with placeholders:

    {{ content }}

During report generation, the placeholder will be replaced by generated
content.

------------------------------------------------------------------------

# Basic Usage

## Insert Text

``` python
container.add_text(text)
```

## Insert Image

``` python
container.add_image(
    "figure.png",
    width_cm=8,
    align="center"
)
```

## Insert Table

``` python
data = [
    ["Name", "Value"],
    ["Example", "123"]
]

container.add_table(data)
```

------------------------------------------------------------------------

# Project Structure

    klword
    │
    ├─ src/
    │   └─ klword/
    │       ├─ __init__.py
    │       ├─ word_api.py
    │       ├─ word_styles.py
    │       └─ container.py
    │
    ├─ demo/
    │   └─ example.py
    │
    ├─ tests/
    │
    ├─ README.md
    ├─ LICENSE
    └─ pyproject.toml

------------------------------------------------------------------------

# Dependencies

This project is built on top of the following libraries:

-   docxtpl
-   python-docx
-   docxcompose
-   jinja2

------------------------------------------------------------------------

# Use Cases

Typical use cases include:

-   automated testing reports
-   simulation result reports
-   batch document generation
-   CI/CD pipeline reports
-   data analysis documentation

------------------------------------------------------------------------

# Contributing

Contributions are welcome.

If you would like to improve this project:

1.  Fork the repository
2.  Create a feature branch
3.  Submit a pull request

------------------------------------------------------------------------

# License

MIT License

Copyright (c) 2026
