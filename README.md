# Installation

    cargo install --git https://github.com/sighol/excel-to-txt

Add this to your .gitconfig:

```gitconfig
[diff "xlsx"]
    textconv = excel-to-txt
```

And add this to your .gitattributes file in the root folder of your repository

```gitattributes
*.xlsx diff=xlsx
```