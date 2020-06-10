# Punchtool

Punchtool is a tool for members in any weixin group to count whether they have studyed today:smile:.

Punchtool 是一个面向微信群成员的工具，用于记录他们今天是否学习:blush:。

## Table of Contents

- [Punchtool](#punchtool)
  - [Table of Contents](#table-of-contents)
  - [Change Log](#change-log)
    - [v0.2(2020/06/10 10:44)](#v0220200610-1044)
    - [v0.12(2020/06/07 23:20:25)](#v01220200607-232025)
    - [v0.11(2020/06/07 09:09:46)](#v01120200607-090946)
    - [v0.10(2020/06/06 15:28:54)](#v01020200606-152854)
    - [v0.1(2020/06/06 14:38:12)](#v0120200606-143812)
  - [Security](#security)
    - [Any optional sections](#any-optional-sections)
  - [Background](#background)
    - [Any optional sections](#any-optional-sections-1)
  - [Install](#install)
    - [Any optional sections](#any-optional-sections-2)
  - [Usage](#usage)
  - [License](#license)


## Change Log

### v0.2(2020/06/10 10:44)
- Biggg Update!:star:
- add: automatically ignore strange emoticons(include emoji) in weixin_name, which means there will be more accurate and matching names!
- fix: You must have two files(**【打卡】名单.xlsx, 【打卡】聊天记录.docx**), it's essential.:D
- add: in **【打卡】名单.xlsx**, I add a column "未打卡次数". if a person have lacked x times, we won't count him.

### v0.12(2020/06/07 23:20:25)
- fix bug: fix the procedure of reading record of chatting.

### v0.11(2020/06/07 09:09:46)
- fix bug: ~word

### v0.10(2020/06/06 15:28:54)
- add: output to xlsx

### v0.1(2020/06/06 14:38:12)
- First commit

## Security

### Any optional sections

## Background

### Any optional sections

## Install

Our project is developd by Python3.6. If you want to run our code, you should install Python.
Then, you need install these following packages.

```shell
pip install python-docx
pip install openpyxl
```

### Any optional sections

## Usage

```
```



Note: The `license` badge image link at the top of this file should be updated with the correct `:user` and `:repo`.

## License

[MIT](../LICENSE)