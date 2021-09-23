# -*- coding: utf-8 -*-
import os
import pandas as pd
import win32com.client as win32
import yaml

from gooey import Gooey, GooeyParser


@Gooey
def main():
    parser = GooeyParser(description='Split excel into multiple files based on column')
    subparsers = parser.add_subparsers(help='sub-command help')
    parser_split = subparsers.add_parser('split', help='split by branch and send mail')
    parser_split.add_argument('input_file', type=str, help='path to the customer_review.xlsx file', widget='FileChooser')

    parser_split.add_argument('-n', '--sheet_name', required=True, help='Sheet name to filter customer')
    parser_split.add_argument('-r', '--skip_rows', type=int, default=3, help='Number of lines to skip at the start of the file (default=%(default)s)')
    parser_split.add_argument('-c', '--column_name', type=str, default='Tên CN', help='Column name to group by (default=%(default)s)')

    parser_split.add_argument('-s', '--mail_subject', required=True, help='Mail subject')
    parser_split.add_argument('-b', '--mail_body', required=True, help='Mail body')
    parser_split.add_argument('--send', default=False, action='store_true', help='When you want to send mail immediately')
    parser_split.set_defaults(func=split)

    parser_merge = subparsers.add_parser('merge', help='merge all received excel files in a directory')
    parser_merge.add_argument('-d', '--directory', help='The directory where you keep all excel files')
    parser_merge.add_argument('output_file', help='The output Excel .xlsx file')
    parser_merge.set_defaults(func=merge)

    args = parser.parse_args()
    args.func(args)


def split(args):
    df = pd.read_excel(args.input_file, sheet_name=args.sheet_name, engine='openpyxl', skiprows=args.skip_rows)
    output_dir = os.path.splitext(args.input_file)[0]
    if not os.path.exists(output_dir):
        os.mkdir(output_dir)
    grouped_df = df.groupby([args.column_name])
    for branch, item in grouped_df:
        grouped_df.get_group(branch).to_excel(os.path.join(output_dir, branch+".xlsx"))

    with open('config.yaml', 'r') as file:
        cfg = yaml.safe_load(file)

    cwd = os.path.dirname(os.path.realpath(__file__))
    attachment_dir = os.path.join(cwd, output_dir)
    for branch, email_addresses in cfg.items():
        send_mail(args.send, attachment_dir, branch, email_addresses, args.mail_subject, args.mail_body)


def send_mail(send, attachment_dir, branch, email_addresses, subject, body):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = "; ".join(email_addresses)
    mail.Subject = subject
    attachment = os.path.join(attachment_dir, branch+".xlsx")
    mail.Attachments.Add(attachment)
    mail.GetInspector

    index = mail.HTMLBody.find('>', mail.HTMLBody.find('<body'))
    mail.HTMLBody = mail.HTMLBody[:index + 1] + body + mail.HTMLBody[index+1:]

    if send:
        mail.Send()
    else:
        mail.Display(False)


def merge(args):
    if args.directory is None:
        args.directory = os.path.abspath('')
    files = os.listdir(args.directory)
    df = pd.DataFrame()
    for file in files:
        if file.endswith('.xlsx'):
            file_df = pd.read_excel(os.path.join(args.directory, file), engine='openpyxl')
            df = df.append(file_df)
    df.to_excel(args.output_file)


if __name__ == "__main__":
    main()