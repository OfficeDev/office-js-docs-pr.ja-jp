# <a name="whats-changed-in-the-javascript-api-for-office"></a>JavaScript API for Office の変更点

JavaScript API for Office は、Office アドインの機能を拡張するため、オブジェクト、メソッド、プロパティ、イベント、列挙体の新規追加や更新によって定期的に更新が加えられています。新規および更新された API のメンバーを確認するには、次のリンクを参照してください。

新しい API メンバーを使用してアドインを開発するには、[プロジェクトで JavaScript API for Office ファイルを更新する](https://docs.microsoft.com/office/dev/add-ins/develop/update-your-javascript-api-for-office-and-manifest-schema-version)必要があります。

前回の更新から変更されていない API メンバーを含むすべての API メンバーを表示するには、「[JavaScript API for Office](javascript-api-for-office.md)」を参照してください。

## <a name="new-and-updated-apis"></a>新規および更新された API

### <a name="new-and-updated-objects"></a>新規オブジェクトと更新されたオブジェクト

|**オブジェクト**|**説明**|**追加または更新されたバージョン**|
|:-----|:-----|:-----|
|`Item`|次に対して更新および追加が行われました。<br><ul><li><p>ユーザーの選択の取得と、メッセージまたは予定の件名と本文を上書きするための、`getSelectedDataAsync` および `setSelectedDataAsync` メソッド。</p></li><li><p>予定の返信フォームへの添付ファイルの追加をサポートする `displayReplyAllForm` および `displayReplyForm` メソッド。</p></li></ul>|Mailbox 1.2|
|`Item`|新規作成モードの Outlook アドインを作成するためのメソッドとフィールドを含めるよう更新されました。 |1.1|
|`Binding`|Access 用コンテンツ アドインにおけるテーブル バインドをサポートするよう更新されました。|1.1|
|`Bindings`|Access 用コンテンツ アドインにおけるテーブル バインドをサポートするよう更新されました。|1.1|
|`Body`|新規作成モードの Outlook アドインでメッセージや予定の本文を作成および編集できるよう追加されました。|1.1|
|`Document`|次に対して更新および追加が行われました。 <ul><li><p>Access 用のコンテンツ アドインで <a href="https://docs.microsoft.com/javascript/api/office/office.document?view=office-js" target="_blank">mode</a>、<a href="https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#settings" target="_blank">settings</a>、および <a href="https://docs.microsoft.com/javascript/api/office/office.document?view=office-js" target="_blank">url</a> の各プロパティをサポートします。</p></li><li><p>PowerPoint および Word 用アドインで、<a href="https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getfileasync-filetype--options--callback-" target="_blank">getFileAsync</a> メソッドを使用してドキュメントを PDF として取得します。</p></li><li><p>Excel、PowerPoint、および Word 用アドインで、<a href="https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getfilepropertiesasync-options--callback-" target="_blank">getFileProperties</a> メソッドを使用してファイルのプロパティを取得します。</p></li><li><p>Excel および PowerPoint 用アドインで、<a href="https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#gotobyidasync-id--gototype--options--callback-" target="_blank">goToByIdAsync</a> メソッドを使用して、ドキュメント内の場所とオブジェクトに移動します。</p></li><li><p>PowerPoint 用アドインで、<a href="https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-" target="_blank">getSelectedDataAsync</a> メソッドを使用して (新しい <span class="keyword">Office.CoercionType.SlideRange</span><a href="https://docs.microsoft.com/javascript/api/office/office.coerciontype?view=office-js" target="_blank">coercionType</a> 列挙体を指定した場合)、選択したスライドの ID、タイトル、およびインデックスを取得します。</p></li></ul>|1.1|
|`Location`|新規作成モードの Outlook アドインで予定の場所を設定できるよう追加されました。|1.1|
|`Office`|Access 用コンテンツ アドインにおけるバインドの取得をサポートするよう select メソッドが更新されました。|1.1|
|`Recipients`|新規作成モードでメッセージや予定の受信者を取得および設定できるよう追加されました。|1.1|
|`Settings`|Access 用コンテンツ アドインにおけるカスタム設定の作成をサポートするよう更新されました。|1.1|
|`Subject`|新規作成モードの Outlook アドインでメッセージや予定の件名を取得および設定できるよう追加されました。|1.1|
|`Time`|新規作成モードの Outlook アドインで予定の開始時刻および終了時刻を取得および設定できるよう追加されました。|1.1|

### <a name="new-and-updated-enumerations"></a>新規列挙体および更新された列挙体

|**オブジェクト**|**説明**|**バージョン**|
|:-----|:-----|:-----|
|`ActiveView`|ユーザーがドキュメントを編集できるかどうかなど、ドキュメントのアクティブなビューの状態を示します。PowerPoint 用アドインで、ユーザーがプレゼンテーション ( **スライド ショー**) を閲覧しているのか、スライドを編集しているのかを判断できるように追加されました。 |1.1|
|`CoercionType`|PowerPoint 用アドインで、**getSelectedDataAsync** メソッドを使用して選択されたスライドの範囲の取得をサポートするよう **Office.CoercionType.SlideRange** が更新されました。|1.1|
|`EventType`|新しい ActiveViewChanged イベントを含めるよう更新されました。|1.1|
|`FileType`|PDF 形式での出力を指定するよう更新されました。|1.1|
|`GoToType`|ドキュメントの移動先の場所またはオブジェクトを指定するよう追加されました。|1.1|

