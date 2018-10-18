# <a name="outlook-add-in-api-requirement-set-17"></a>Outlook アドイン API 要件セット 1.7

JavaScript API for Office の Outlook アドイン API サブセットには、Outlook アドインで利用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。

## <a name="whats-new-in-17"></a>1.7 の新機能

要件セット 1.7 には、 [要件セット 1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) のすべての機能が含まれています。 次の機能を追加しました。

- 打ち合わせと、会議出席依頼用メッセージの定期的パターンに関する新しい Api を追加しました。
- item.をプロパティから作成モードでも利用できるように変更しました。
- RecurrenceChanged、RecipientsChanged、AppointmentTimeChanged イベントのサポートを追加しました。

### <a name="change-log"></a>変更ログ

- 「[From](/javascript/api/outlook_1_7/office.from)」の追加：From値を取得するメソッドを提示する新しいオブジェクトを追加します。
- 「 [開催者](/javascript/api/outlook_1_7/office.organizer)」の追加：開催者の値を取得するメソッドを提示する新しいオブジェクトを追加します。
- 「 [定期的なアイテム](/javascript/api/outlook_1_7/office.recurrence)」の追加： 打ち合わせの定期的パターンと会議以来に限ったメッセージの定期的パターンを取得し設定するためのメソッドを提示する新しいオブジェクトを追加します。
- 「 [RecurrenceTimeZone](/javascript/api/outlook_1_7/office.recurrencetimezone)」の追加：定期的パターンのタイム ゾーンの構成を表す新しいオブジェクトを追加します。
- 「 [SeriesTime](/javascript/api/outlook_1_7/office.seriestime)」の追加:：定期的一連の打ち合わせの日時を取得し設定するメソッドと、一連の定期的会議出席依頼の日時を取得するメソッドを提示する新しいオブジェクトを追加します。
- 「[Office.context.mailbox.addHandlerAsync](office.context.mailbox.item.md#addhandlerasynceventtype-handler-options-callback) 」の追加： サポートするイベント用のイベント ハンドラーを加えるための新しいメソッドを追加します。
- 「 [Office.context.mailbox.item.from](office.context.mailbox.item.md#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom)」の変更：作成モードのfrom値を取得するための変更をします。
- 「[Office.context.mailbox.item.organizer](office.context.mailbox.item.md#organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer)」の変更：作成モードの開催者値を取得するための変更をします。
- 「 [Office.context.mailbox.item.recurrence](office.context.mailbox.item.md#nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence)」の追加：打ち合わせアイテムの定期的パターンを管理するためのメソッドを提示するオブジェクトを取得し設定する新しいプロパティを追加します。 このプロパティを使用して、会議出席依頼アイテムの定期的なパターンを取得することもできます。
- 「 [Office.context.mailbox.item.removeHandlerAsync](office.context.mailbox.item.md#removehandlerasynceventtype-handler-options-callback)」の追加：イベント ハンドラーを削除する新しいメソッドを追加します。
- 「 [Office.context.mailbox.item.seriesId](office.context.mailbox.item.md#nullable-seriesid-string)」の追加：パターンのある一連アイテムのIDを取得するための新しいプロパティを追加します。
- 「 [Office.MailboxEnums.Days](/javascript/api/outlook_1_7/office.mailboxenums.days)」の追加：  週の曜日または日の種類を指定する新しい列挙型を追加します。
- 「 [Office.MailboxEnums.Month](/javascript/api/outlook_1_7/office.mailboxenums.month)」の追加：月を指定する新しい列挙型を追加します。
- 「 [Office.MailboxEnums.RecurrenceTimeZone](/javascript/api/outlook_1_7/office.mailboxenums.recurrencetimezone)」の追加：パターンに適用するタイム ゾーンを指定する新しい列挙型を追加します。
- 「 [Office.MailboxEnums.RecurrenceType](/javascript/api/outlook_1_7/office.mailboxenums.recurrencetype)」の追加： パターンの種類を指定する新しい列挙型を追加します。
- 「 [Office.MailboxEnums.WeekNumber](/javascript/api/outlook_1_7/office.mailboxenums.weeknumber)」の追加：その月の週を指定する新しい列挙型を追加します。
- 「[Office.EventType](/javascript/api/office/office.eventtype)」の変更：RecurrenceChanged、RecipientsChanged、および AppointmentTimeChanged のイベントを、 `RecurrenceChanged`、 `RecipientsChanged`、 `AppointmentTimeChanged` 、それぞれのエントリの追加によってサポートするための変更をします。

## <a name="see-also"></a>関連項目

- [Outlook アドイン](https://docs.microsoft.com/outlook/add-ins/)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [作業の開始](https://docs.microsoft.com/outlook/add-ins/quick-start)