# <a name="outlook-add-in-api-requirement-set-16"></a>Outlook アドイン API 要件セット 1.6

JavaScript API for Office の Outlook アドイン API サブセットには、Outlook アドインで利用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。

> [!NOTE]
> このドキュメントは、最新の要件セット以外の[要件セット](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)向けです。

## <a name="whats-new-in-16"></a>1.6 の最新情報

要件セット 1.6 には、[要件セット 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) のすべての機能が含まれています。 次の機能を追加しました。

- ユーザーがアドインを有効にするために選択したエンティティまたは RegEx を取得する 文脈アドインの新しい API を追加しました。
- 新しいメッセージ フォームを開く新しい API を追加しました。
- アドインがユーザーのメールボックスのアカウントの種類を決定するための機能を追加しました。

### <a name="change-log"></a>変更ログ

- [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#getselectedentities--entitiesjavascriptapioutlook16officeentities): ユーザーが選択した強調表示された一致内で見つかったエンティティを取得する新機能を追加します。 強調表示された一致は、コンテキスト アドインに適用されます。
- [ Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#getselectedregexmatches--object)  : マニフェスト XML ファイルで定義した正規表現と一致する、強調表示された一致内の文字列値を返す新機能を追加します。 強調表示された一致は、コンテキスト アドインに適用されます。
- [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#displaynewmessageformparameters)を追加: 新しいメッセージ フォームを表示する新しい関数を追加します。
- [Office.context.mailbox.userProfile.accountType](office.context.mailbox.userprofile.md#accounttype-string)を追加: ユーザーのアカウントの種類を示すユーザー プロファイルに新しいメンバーを追加します。

## <a name="see-also"></a>関連項目

- [Outlook アドイン](https://docs.microsoft.com/outlook/add-ins/)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [作業の開始](https://docs.microsoft.com/outlook/add-ins/quick-start)