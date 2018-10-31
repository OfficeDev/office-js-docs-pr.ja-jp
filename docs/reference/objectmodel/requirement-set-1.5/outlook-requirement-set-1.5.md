# <a name="outlook-add-in-api-requirement-set-15"></a>Outlook アドイン API 要件セット 1.5

JavaScript API for Office の Outlook アドイン API サブセットには、Outlook アドインで利用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。

> [!NOTE]
> このドキュメントは、最新の要件セット以外の[要件セット](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)が対象です。

## <a name="whats-new-in-15"></a>1.5 の新機能

要件セット 1.5 は、[要件セット 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md) のすべての機能を含み、次の機能を追加しています。

- [ピン留め可能な作業ウィンドウ](https://docs.microsoft.com/outlook/add-ins/pinnable-taskpane)のサポートを追加しました。
- [REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api) 呼び出しのサポートを追加しました。
- インラインで添付ファイルにマークを付ける機能を追加しました。
- 作業ウィンドウまたはダイアログを閉じる機能を追加しました。

### <a name="change-log"></a>変更ログ

- [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#addhandlerasynceventtype-handler-options-callback) を追加: サポートするイベントのイベント ハンドラーを追加します。
-  [Office.EventType](office.md#eventtype-string)を追加: イベント ハンドラーに関連付けられているイベントを指定し、ItemChanged イベントのサポートが含みます。
- [Office.context.mailbox.restUrl](office.context.mailbox.md#resturl-string) を追加: 電子メール アカウントの REST エンドポイントの URL を取得します。
- [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#getcallbacktokenasyncoptions-callback) を変更: このメソッドの新しい署名の付いた新しいバージョン (`getCallbackTokenAsync([options], callback)`) を追加しました。変更のない元のバージョンも引き続き使用できます。
- [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--) を追加しました。
- [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback) を変更: `isInline` と呼ばれる `options` ディクショナリの新しい値です。メッセージ本文でイメージをインラインで使用することを指定するために使用します。
- [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata) を変更: `isInline` と呼ばれる `formData.attachments` ディクショナリの新しい値です。メッセージ本文でイメージをインラインで使用することを指定するために使用します。
- [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata)  を変更: `isInline` と呼ばれる `formData.attachments` ディクショナリの新しい値です。メッセージ本文でイメージをインラインで使用することを指定するために使用します。

## <a name="see-also"></a>関連項目

- [Outlook アドイン](https://docs.microsoft.com/outlook/add-ins/)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [はじめましょう](https://docs.microsoft.com/outlook/add-ins/quick-start)