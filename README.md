# xsplitter
Group customer by branch and send mail using Outlook.

## Copy to the Windows machine

```
$webClient = New-Object System.Net.WebClient
$webClient.Proxy = [System.Net.GlobalProxySelection]::GetEmptyWebProxy()
$webClient.DownloadFile("http://192.168.1.88:8000/main.exe", "main.exe")
```

## Usage
- Open the Command Prompt by typing `cmd` in the Search box
- Change to the `xsplitter` directory:

```
> cd xsplitter
```

- List files and folders in the current directory:

```
> dir
```

- How to use this program?

```
> main.exe -h
usage: main.exe [-h] -n SHEET_NAME [-r SKIP_ROWS] [-c COLUMN_NAME] -s MAIL_SUBJECT -b MAIL_BODY review_file

Split excel into multiple files based on column

positional arguments:
  review_file           path to the customer_review.xlsx file

optional arguments:
  -h, --help            show this help message and exit
  -n SHEET_NAME, --sheet_name SHEET_NAME
                        Sheet name to filter customer
  -r SKIP_ROWS, --skip_rows SKIP_ROWS
                        Number of lines to skip at the start of the file
  -c COLUMN_NAME, --column_name COLUMN_NAME
                        Column name to group by
  -s MAIL_SUBJECT, --mail_subject MAIL_SUBJECT
                        Mail subject
  -b MAIL_BODY, --mail_body MAIL_BODY
                        Mail body
```

- Please note that file, sheet name, mail subject, mail body are required:

```
> main.exe
usage: main.exe [-h] -n SHEET_NAME [-r SKIP_ROWS] [-c COLUMN_NAME] -s MAIL_SUBJECT -b MAIL_BODY review_file
main.py: error: the following arguments are required: review_file, -n/--sheet_name, -s/--mail_subject, -b/--mail_body
```

- This program will group by branch and send mail immediately, please make sure that your config file is correct:

```
AGG:
  - email1@abc.com.vn
BLI:
  - email2@abc.com.vn
```

- Copy your file to the `xsplitter` directory, note down the sheet name and run:

```
> main.exe "file.xlsx" -n "PS giảm" -s "Rà soát KH XNK" -b "Kính đề nghị ... <br><br> Trân trọng cảm ơn."
```

In the mail body, whenever you want to insert a line break, just use `<br>`.