import datetime
import zoneinfo
import os.path
import argparse
from icecream import ic
import os

from hinocal import util


def main2(command=None, in_file=None, args=None):

    if command == "list":
        events = get_events(service, args.startdate)

        if not events:
            print("No upcoming events found.")
            return

        # Prints the start and name of the next 10 events
        for event in events:
            start = event["start"].get("dateTime", event["start"].get("date"))
            # print(start, event["id"], event["summary"], event["description"])
            ic(event)
            # created_str = event.get("created")
            # created = iso2jst(created_str)
            # updated_str = event.get("updated")
            # updated = iso2jst(updated_str)
            # ic(event.get("summary"))
            # ic(created_str, created.strftime("%Y-%m-%d %H:%M:%S"))
            # ic(updated_str, updated.strftime("%Y-%m-%d %H:%M:%S"))

    if command == "calendar":
        list_calendar(service)

    if command == "download":
        download_events(service, args.school_year, args.calendar_file)

    if command == "upload":
        upload_events(service, args.school_year, args.calendar_file)

    if command == "sync":
        ic(in_file)
        if in_file == None:
            print("Please specify the file name for input.")
            return False
        wb = openpyxl.load_workbook(in_file)
        ws = wb["Sheet1"]
        max_row = ws.max_row
        counter = 0
        for row in ws:
            counter += 1
            # if counter > 10:  # debug
            #     break
            if row[0].value == "日付":
                # skip
                print(f"{counter}/{max_row}: skip title row")
                pass
            else:
                g_event = create_event_from_row(row)
                if g_event:
                    event = update_event(service, g_event)
                    if event:
                        print(
                            f"{counter}/{max_row}: Updated. {event["start"]["dateTime"]} - {event["end"]["dateTime"]}: {event["summary"]} {event["description"]}"
                        )
                        # write back to excel sheet
                        ws.cell(counter, 4, g_event["id"])
                    else:
                        print(
                            f"{counter}/{max_row}: Skipped. {g_event["start"]["dateTime"]} - {g_event["end"]["dateTime"]}: {g_event["summary"]}"
                        )
                else:
                    # ignore None
                    pass
        wb.save(in_file)

        try:
            pass

            # event = update_event(service, event)
            # print ('Event created: %s' % (event.get('htmlLink')))

        except HttpError as error:
            print(f"An error occurred: {error}")


def list(args, service):
    """list command"""
    events = util.get_events(service, args.year_month)

    if not events:
        print("No upcoming events found.")
        return

    # Prints the start and name of the next 10 events
    for event in events:
        if args.detail:
            ic(event)
        else:
            start = event["start"].get("dateTime", event["start"].get("date"))
            description = util.remove_time_stamp(event.get("description"))
            print(start, event["id"], event["summary"], description)


def calendar(args, service):
    """calendar command"""
    util.list_calendar(service)


def download(args, service):
    """download command"""
    util.download_events(service, args.school_year, args.calendar_file)


def upload(args, service):
    """upload command"""
    util.upload_events(service, args.school_year, args.calendar_file)


def main():
    parser = argparse.ArgumentParser(
        prog="hinocal",
        description="Googleカレンダーのイベント(予定)を登録・更新する。 Googleカレンダーの情報を正と捉え、更新するためにダウンロードし、編集後にアップロードしてカレンダーを更新する。 対象とするカレンダーは.envファイルの`CAL_ID`に記入すること。",
    )
    subparsers = parser.add_subparsers(help="sub-command help", required=True)

    # list
    parser_list = subparsers.add_parser(
        "list",
        help="カレンダーのイベントの一部を表示する",
        description="Googleカレンダーからイベントを取得して表示する",
    )
    parser_list.set_defaults(func=list)
    parser_list.add_argument(
        "-mo", "--year_month", help="yyyy-mm 処理対象とする年月を指定する。"
    )
    parser_list.add_argument(
        "-vv", "--detail", action="store_true", help="イベント全体を表示する"
    )

    # calendar
    parser_cal = subparsers.add_parser(
        "calendar",
        help="カレンダー一覧を表示する",
        description="サインインしているユーザのカレンダー一覧を表示する",
    )
    parser_cal.set_defaults(func=calendar)

    # download
    parser_download = subparsers.add_parser(
        "download",
        help="カレンダーのイベントを取得しExcelファイルに記入/作成する",
        description="カレンダーからイベントを取得しExcelファイルに記入する。 取得範囲は年度単位。何も指定しないと本年度のイベント一覧を取得する。",
    )
    parser_download.add_argument(
        "-sy", "--school_year", help="年度 yyyy。処理対象とする年度を指定する。"
    )
    parser_download.add_argument(
        "-cf",
        "--calendar_file",
        help="カレンダーの内容を書き出すexcelファイル名。 指定しない場合は'calendar_syYYYY.xlsx'。 既存のファイルは上書きされる。 通常は、年度(--school_year)を指定し、ファイル名は指定しなくてよい。",
    )
    parser_download.set_defaults(func=download)

    # upload
    parser_upload = subparsers.add_parser(
        "upload",
        help="Excelに記載されたイベントをカレンダーに登録/更新/削除する",
        description="Excelに記載されたイベントをカレンダーに登録する。 downloadで取得したイベントは修正される。新規追加した行は登録される。 downloadで取得したイベントの開始日と行事を空白にするとそのイベントはカレンダーから消去される。 Excelを修正する際は`イベントID`を修正しないよう注意すること。",
    )
    parser_upload.add_argument(
        "-sy", "--school_year", help="年度 yyyy。処理対象とする年度を指定する。"
    )
    parser_upload.add_argument(
        "-cf",
        "--calendar_file",
        help="カレンダーの内容を読み込むexcelファイル名。 指定しない場合は'calendar_syYYYY.xlsx'。 通常は、年度(--school_year)を指定し、ファイル名は指定しなくてよい。",
    )
    parser_upload.set_defaults(func=upload)
    parser.add_argument(
        "-re",
        "--relogin",
        action="store_true",
        help="サインイン情報をクリアしてから実行する",
    )
    # parser.add_argument("-sd", "--startdate", help="開始年月 yyyy-mm")
    # parser.add_argument("-f", "--file", help="行事予定一覧excelファイル")
    # parser.add_argument(
    #     "-cf",
    #     "--calendar_file",
    #     help="カレンダーの内容を書き出す/読み込むexcelファイル名。指定しない場合は'calendar_syYYYY.xlsx'。既存のファイルは上書きされる。",
    # )
    # parser.add_argument(
    #     "-sy", "--school_year", help="年度 yyyy。downloadとuploadの際に使用する。"
    # )
    args = parser.parse_args()

    if args.relogin:
        if os.path.exists("./token.json"):
            os.remove("./token.json")

    service = util.get_service()

    args.func(args, service)

    # main2(args.command, args.file, args)


if __name__ == "__main__":
    main()
