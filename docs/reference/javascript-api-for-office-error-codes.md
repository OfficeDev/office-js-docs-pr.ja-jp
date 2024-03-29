---
title: Office の一般的な API エラー コード
description: この記事では、Office Common API の使用中に発生する可能性があるエラー メッセージについて説明します。
ms.date: 09/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: d77b4c0c458e11da0057f06a5088ef8a28e4ccd2
ms.sourcegitcommit: 05be1086deb2527c6c6ff3eafcef9d7ed90922ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/28/2022
ms.locfileid: "68092981"
---
# <a name="office-common-api-error-codes"></a>Office の一般的な API エラー コード

この記事では、Common API モデルの使用中に発生する可能性があるエラー メッセージについて説明します。 これらのエラー コードは、Excel JavaScript API や Word JavaScript API などのアプリケーション固有の API には適用されません。

共通 [API とアプリケーション固有の API モデル](../develop/understanding-the-javascript-api-for-office.md#api-models) の違いの詳細については、API モデルを参照してください。

## <a name="error-codes"></a>エラー コード

次の表に、エラー コード、名前、表示されるメッセージ、それらが示す状態を示します。

|Error.code|Error.name|Error.message|条件|
|:-----|:-----|:-----|:-----|
|1000|無効な強制型変換|指定された強制型変換はサポートされていません。|強制型は、Office アプリケーションではサポートされていません。 (たとえば、OOXML 型と HTML 強制型は Excel ではサポートされていません。|
|1001|データの読み取りエラー|現在の選択項目はサポートされていません。|ユーザーの現在の選択項目はサポートされていません (つまり、サポートされている強制型変換と異なっている部分があります)。|
|1002|無効な強制型変換|指定された強制型変換は、このバインド タイプと互換性がありません。|ソリューション開発者が指定した強制型変換とバインド タイプの組み合わせには互換性がありません。|
|1003|データの読み取りエラー|指定した rowCount または columnCount の値が無効です。|ユーザーが無効な列数または行数を指定しています。|
|1004|データの読み取りエラー|現在の選択項目は指定された強制型変換と互換性がありません。|このアプリケーションでは、現在の選択項目は指定された強制型変換でサポートされていません。|
|1005|データの読み取りエラー|指定された startRow または startColumn の値が正しくありません。|ユーザーが無効な startRow または startCol の値を指定しています。|
|1006|データの読み取りエラー|テーブルに結合されたセルが含まれている場合、座標パラメーターを強制型変換タイプ "Table" と共に使用できません。|ユーザーは一様でないテーブル (つまり、マージされたセルを持つテーブル) から一部のデータを取得しようとしています。 |
|1007|データの読み取りエラー|ドキュメントのサイズが大きすぎます。|ユーザーが、現在サポートされているサイズより大きいドキュメントを取得しようとしています。|
|1008|データの読み取りエラー|要求されたデータ セットが大きすぎます。|ユーザーは、Office アプリケーションによって定義されたデータ制限を超えてデータの読み取りを要求します。|
|1009|データの読み取りエラー|指定されたファイルの種類はサポートされていません。|ユーザーが、無効なファイルの種類を送信しています。|
|2000|データの書き込みエラー|指定されたデータ オブジェクトの型はサポートされていません。 |サポートされていないデータ オブジェクトが指定されています。|
|2001|データの書き込みエラー|現在の選択項目を書き込むことができません。|The user's current selection is not supported for a write operation. (For example, when the user selects an image.)|
|2002|データの書き込みエラー|指定されたデータ オブジェクトは、現在の選択項目の形状または次元と互換性がありません。|複数のセルが選択されています (また、選択項目の形状がデータの形状と一致しません)。 複数のセルが選択されています (また、選択項目の次元がデータの次元と一致しません)。|
|2003|データの書き込みエラー|指定されたデータ オブジェクトがデータを上書きするため、設定操作に失敗しました。|1 つのセルが選択され、指定されたデータ オブジェクトが、ワークシート内のデータを上書きします。|
|2004|データの書き込みエラー|指定されたデータ オブジェクトが、現在の選択項目のサイズと一致しません。|ユーザーが、現在の選択項目のサイズよりも大きいオブジェクトを指定しています。|
|2005|データの書き込みエラー|指定された startRow または startColumn の値が正しくありません。|ユーザーが無効な startRow または startCol の値を指定しています。|
|2006|無効な形式のエラー|指定されたデータ オブジェクトの形式が正しくありません。|ソリューション開発者が、HTML または OOXML の無効な文字列、HTML の不正な文字列、または OOXML の無効な文字列を指定しています。|
|2007|無効なデータ オブジェクト|指定されたデータ オブジェクトの型は、現在の選択項目と互換性がありません。|ソリューション開発者が、指定された強制型変換と互換性のないデータ オブジェクトを指定しています。|
|2008|データの書き込みエラー|TBD|TBD|
|2009|データの書き込みエラー|指定されたデータ オブジェクトが大きすぎます。|ユーザーは、Office アプリケーションによって定義されたデータ制限を超えてデータを設定しようとします。|
|2010|データの書き込みエラー|テーブルに結合されたセルが含まれている場合は、座標パラメーターを強制変換タイプ Table と共に使用できません。|ユーザーが一様でないテーブル (つまり、マージされたセルを持つテーブル) から一部のデータを設定しようとしています。|
|3000|バインディングの作成エラー|現在の選択項目をバインドできません。|The user's selection is not supported for binding. (For example, the user is selecting an image or other non-supported object.)|
|3001|バインディングの作成エラー|TBD|TBD|
|3002|無効なバインド エラー|指定されたバインドが存在しません。|開発者は、存在しない、または削除されたバインディングにバインドしようとしています。|
|3003|バインディングの作成エラー|連続していない選択項目はサポートされません。|ユーザーが複数の選択を行っています。|
|3004|バインディングの作成エラー|現在の選択項目と指定されたバインド タイプでバインドを作成できません。|There are several conditions under which this might happen. Please see the "Binding creation error conditions" section later in this article.|
|3005|無効なバインド操作|このバインド タイプではサポートされていない操作です。|開発者は、強制型ではないバインド型に対して行の追加または列の追加操作を送信します `table`。|
|3006|バインディングの作成エラー|名前付きアイテムが存在しません。|The named item cannot be found. No content control or table with that name exists.|
|3007|バインディングの作成エラー|同じ名前を持つ複数のオブジェクトが見つかりました。|競合エラー: 同じ名前のコンテンツ コントロールが複数存在し、競合時に失敗が設定 `true`されています。|
|3008|バインディングの作成エラー|指定されたバインド タイプは、指定された名前付きアイテムと互換性がありません。|名前付きアイテムを型にバインドすることはできません。 たとえば、コンテンツ コントロールにはテキストが含まれていますが、開発者は強制型 `table`を使用してバインドしようとしました。|
|3009|無効なバインド操作|バインド タイプがサポートされていません。|下位互換性のために使用されます。|
|3010|サポートされないバインド操作|選択するコンテンツはテーブル形式にする必要があります。 データをテーブルとして書式設定して、もう一度やり直してください。|開発者は、強制型`matrix`のデータに対して`TableBinding`オブジェクトのメソッドを`deleteAllDataValuesAsync`使用`addRowsAsync`しようとしています。|
|4000|設定の読み取りエラー|指定された設定の名前が存在しません。|存在しない設定の名前が指定されています。|
|4001|設定の保存エラー|設定を保存できませんでした。|設定を保存できませんでした。|
|4002|古い設定のエラー|設定が古いために保存できませんでした。|設定が古く、開発者が設定を上書きしないよう指定しています。|
|5000|古い設定のエラー|この操作はサポートされていません。|この操作は、現在の Office アプリケーションではサポートされていません。 たとえば、 `document.getSelectionAsync` Outlook から呼び出されます。|
|5001|内部エラー|内部エラーが発生しました。|内部エラーが発生しています。これは、次のいずれかの理由で発生します。<br/><table><tr><td>ブックを共有している他のユーザーが使用しているアドインが、ほとんど同時にバインドを作成しました。使用しているアドインは、再バインドを行う必要があります。</tr></td><tr><td>不明なエラーが発生しました。</tr></td><tr><td>処理に失敗しました。</tr></td><tr><td>ユーザーが権限を持つロールのメンバーではないために、アクセスが拒否されました。</tr></td><tr><td>セキュリティで保護された、暗号化された通信が必要なために、アクセスが拒否されました。</tr></td><tr><td>データが古いので、クエリがデータを再取得できるよう確認する必要があります。</tr></td><tr><td>サイト コレクションの CPU クォータが限界を超えています。</tr></td><tr><td>サイト コレクションのメモリ クォータが限界を超えています。</tr></td><tr><td>セッションのメモリ クォータが限界を超えています。</tr></td><tr><td>ブックが無効な状態なので、操作を実行できません。</tr></td><tr><td>アイドル状態が続いてセッションがタイムアウトしました。ユーザーがブックを再読み込みする必要があります。</tr></td><tr><td>ユーザーごとに許可されるセッションの最大数を超えています。</tr></td><tr><td>操作はユーザーによって取り消されました。</tr></td><tr><td>時間がかかりすぎているため、操作を完了できません。</tr></td><tr><td>要求を完了できません。再試行する必要があります。</tr></td><tr><td>製品の試用期間の期限が切れています。</tr></td><tr><td>アイドル状態が続いたのでセッションがタイムアウトしました。</tr></td><tr><td>ユーザーは指定されたセル範囲に対する操作を実行する権限がありません。</tr></td><tr><td>現在のコラボレーションのセッションとユーザーの地域の設定が一致しません。</tr></td><tr><td>ユーザーはもはや接続されていません。ブックを更新し再度開く必要があります。</tr></td><tr><td>要求した範囲がシートに存在しません。</tr></td><tr><td>ユーザーは、ブックを編集する権限がありません。</tr></td><tr><td>ブックはロックされているので、編集できません。</tr></td><tr><td>セッションは、ブックを自動的に保存できません。</tr></td><tr><td>セッションは、ブック ファイルのロックを更新できません。</tr></td><tr><td>要求を処理できません。再試行する必要があります。</tr></td><tr><td>ユーザーのサインイン情報を検証できませんでした。再入力する必要があります。</tr></td><tr><td>ユーザーのアクセスが拒否されています。</tr></td><tr><td>共有ブックを更新する必要があります。</tr></td></table>|
|5002|アクセスが拒否されました|要求された操作は、現在のドキュメント モードでは許可されません。|ソリューション開発者が設定操作を送信しましたが、ドキュメントが "編集の制限" など、変更を許可しないモードになっています。|
|5003|イベント登録エラー|指定されたイベントの種類は、現在のオブジェクトではサポートされていません。|ソリューション開発者が、存在しないイベントにハンドラーを登録または登録解除しようとしています。|
|5004|無効な API 呼び出し|現在のコンテキストで無効な API 呼び出しです。|Excel でオブジェクトを使用しようとすると、コンテキストに対して無効な呼び出しが `CustomXMLPart` 行われます。|
|5005|データが古い|サーバー上のデータが古いため、操作が失敗しました。|サーバー上のデータを更新する必要があります。|
|5006|セッションのタイムアウト|ドキュメント セッションがタイムアウトしました。 ドキュメントを再読み込みします。 |セッションがタイムアウトになりました。|
|5007|無効な API 呼び出し|列挙体は、現在のコンテキストではサポートされていません。|列挙体は、現在のコンテキストではサポートされていません。|
|5009|アクセスが拒否されました|アクセスが拒否されました|アドインに特定の API を呼び出すためのアクセス許可がありません。|
|5012|無効またはタイム アウトになるセッション|Office ブラウザー セッションの有効期限が切れているか、無効です。 続行するには、ページを更新します。|Office クライアントとサーバー間のセッションの有効期限が切れたか、お使いのコンピューターで、日付、時刻、タイム ゾーンが正しくありません。|
|6000|無効なノード|指定されたノードが見つかりませんでした。|`CustomXmlPart`ノードが見つかりませんでした。|
|6100|カスタム XML エラー|カスタム XML エラー|無効な API 呼び出し。|
|7000|無効な ID|指定された ID が存在しません。|無効な ID。|
|7001|無効なナビゲーション|ナビゲーションがサポートされていない場所にオブジェクトがあります。|The user can find the object, but cannot navigate to it. (For example, in Word, the binding is to the header, footer, or a comment.)|
|7002|無効なナビゲーション|オブジェクトがロックされているか、保護されています。|ロックまたは保護された範囲へ移動しようとしています。|
|7004|無効なナビゲーション|インデックスが範囲を超えているため、操作に失敗しました。|範囲外のインデックスに移動しようとしています。|
|8000|パラメーターがありません|We couldn't format the table cell because some parameter values are missing. Double-check the parameters and try again.|The cellFormat method is missing some parameters. For example, there are missing cells, format, or tableOptions parameters.|
|8010|無効な値|One or more of the cells parameters have values that aren't allowed. Double-check the values and try again.|The common cells reference enumeration is not defined. For example, All, Data, Headers.|
|8011|無効な値|One or more of the tableOptions parameters have values that aren't allowed. Double-check the values and try again.|tableOptions の値のいずれかが無効です。|
|8012|無効な値|One or more of the format parameters have values that aren't allowed. Double-check the values and try again.|foramt の値のいずれかが正しくありません。|
|8020|範囲外|The row index value is out of the allowed range. Use a positive value (0 or higher) that's less than the number of rows.|行のインデックスが、テーブルの最大行のインデックスより大きいか、または 0 より小さいです。|
|8021|範囲外|The column index value is out of the allowed range. Use a positive value (0 or higher) that's less than the number of columns.|列のインデックスが、テーブルの最大列のインデックスより大きいか、または 0 より小さいです。|
|8022|範囲外|値が許容範囲外です。|形式の値の一部がサポート範囲外です。|
|9016|アクセス許可が拒否されました|アクセスが拒否されました|アクセスが拒否されました。|
|9020|汎用応答エラー|内部エラーが発生しました。|内部エラー状態を参照します。これは、さまざまな理由で発生する可能性があります。|
|9021|エラーの保存|サーバーにアイテムを保存しようとしたときに接続エラーが発生しました。|アイテムを保存できませんでした。 これは、Outlook デスクトップでオンライン モードを使用している場合や、Exchange サーバーから削除された下書きアイテムを再保存しようとした場合にサーバー接続エラーが発生した可能性があります。|
|9022|別のストア エラーのメッセージ|メッセージが別のストアに保存されているため、EWS ID を取得できません。|メッセージが移動されたか、送信メールボックスが変更された可能性があるため、現在のメッセージの EWS ID を取得できませんでした。|
|9041|ネットワーク エラー|ユーザーはネットワークに接続されていません。 ネットワーク接続を確認し、やり直してください。|ユーザーがネットワークまたはインターネットにアクセスできなくなりました。|
|9043|添付ファイルの種類がサポートされていません|添付ファイルの種類はサポートされていません。|API は添付ファイルの種類をサポートしていません。 たとえば、 `item.getAttachmentContentAsync` 添付ファイルがリッチ テキスト形式の埋め込み画像である場合、または電子メールや予定表アイテム以外のアイテムの種類 (連絡先やタスク アイテムなど) の場合、このエラーがスローされます。|
|12002|*適用されません。*|*該当なし。*|以下のいずれか:<br> - `displayDialogAsync` に渡された URL にページが存在しない。<br> - `displayDialogAsync` に渡されたページが読み込まれたが、ダイアログ ボックスが見つからないか読み込むことができないページを指していたか、またはダイアログ ボックスが無効な構文を含む URL を指している。 ダイアログ内でスローされ、ホスト ページの `DialogEventReceived` イベントをトリガーします。|
|12003|*適用されません。*|*適用されません。*|ダイアログ ボックスが HTTP プロトコルを使用している URL を指していました。 HTTPS が必要です。 ダイアログ内でスローされ、ホスト ページの `DialogEventReceived` イベントをトリガーします。|
|12004|*適用されません。*|*該当なし。*|`displayDialogAsync` に渡される URL のドメインは信頼されていません。 ドメインは、ホスト ページと同じドメインにある必要があります (プロトコルとポート番号を含む)。 `displayDialogAsync` の呼び出しによってスローされます。|
|12005|*適用されません。*|*該当なし。*|`displayDialogAsync` に渡される URL には HTTP プロトコルを使用します。 HTTPS が必要です。 `displayDialogAsync` の呼び出しによってスローされます  (Office の一部のバージョンでは、12004 で返されるのと同じエラー メッセージが、12005 でも返されます)。|
|12006|*該当なし。*|*該当なし。*|ダイアログ ボックスが閉じられました。通常は、ユーザーが **X** ボタンを選択したためです。 ダイアログ内でスローされ、ホスト ページの `DialogEventReceived` イベントをトリガーします。|
|12007|*適用されません。*|*適用されません。*|ダイアログ ボックスは、このホスト ウィンドウで既に開いています。 作業ウィンドウなどのホスト ウィンドウで、一度に開けるダイアログ ボックスは 1 つだけです。 `displayDialogAsync` の呼び出しによってスローされます。|
|12009|*適用されません。*|*適用されません。*|ダイアログ ボックスを無視するようにユーザーが選択しました。 このエラーは、ダイアログの表示をアドインに許可しないようにユーザーが選択すると、Office のオンライン バージョンで発生することがあります。 `displayDialogAsync` の呼び出しによってスローされます。|
|12011|*適用されません。*|*適用されません。*|ユーザーのブラウザーは、ポップアップをブロックする方法で構成されます。 このエラーは、ブラウザーが Safari で、ポップアップをブロックするように構成されているか、ブラウザーが Edge Legacy で、アドイン ドメインがダイアログが開こうとしているドメインとは別のセキュリティ ゾーンにある場合、Office on the webで発生する可能性があります。 `displayDialogAsync` の呼び出しによってスローされます。|
|13nnn|*適用されません。*|*適用されません。*|[getAccessToken からのエラーの原因と処理に関する](../develop/troubleshoot-sso-in-office-add-ins.md#causes-and-handling-of-errors-from-getaccesstoken)ページを参照してください。|

## <a name="binding-creation-error-conditions"></a>バインディングの作成エラーの条件

When a binding is created in the API, indicate the binding type that you want to use. The following tables lists the binding types and the resulting binding behaviors that are expected.

### <a name="behavior-in-excel"></a>Excel での動作

次の表は、Excel のバインドの動作をまとめたものです。

|指定されたバインド タイプ|実際の選択|動作|
|:-----|:-----|:-----|
|Matrix|セルの範囲 (1 つのテーブル内にあり、単一セルの場合を含む)|選択したセルに型 `matrix` のバインドが作成されます。 文書内の変更はありません。|
|マトリックス|セル内の選択されたテキスト|セル全体に型 `matrix` のバインドが作成されます。 文書内の変更はありません。|
|マトリックス|複数の選択項目または無効な選択項目 (たとえば、ユーザーが画像、オブジェクト、ワードアートを選択した場合)。|バインドを作成できません。|
|テーブル|セルの範囲 (単一セルの場合を含む)|バインドを作成できません。|
|テーブル|テーブル内のセルの範囲 (テーブル内の 1 つのセル、テーブル全体、またはテーブルのセル内のテキストの場合を含む)。|バインドがテーブル全体で作成されます。|
|テーブル|選択項目の半分がテーブルで、半分がテーブル外です|バインドを作成できません。|
|テーブル|(テーブル内ではなく) セル内で選択されたテキスト。|バインドを作成できません。|
|テーブル|複数の選択項目または無効な選択項目 (たとえば、ユーザーが画像、オブジェクト、ワードアートなどを選択した場合)。|バインドを作成できません。|
|テキスト|セルの範囲|バインドを作成できません。|
|テキスト|テーブル内のセルの範囲|バインドを作成できません。|
|テキスト|1 つのセル|型 `text` のバインドが作成されます。|
|テキスト|テーブル内の 1 つのセル|型 `text` のバインドが作成されます。|
|テキスト|セル内の選択されたテキスト|セル全体に型 `text` のバインドが作成されます。|

### <a name="behavior-in-word"></a>Word の動作

次の表は、Word のバインドの動作をまとめたものです。

|指定されたバインド タイプ|実際の選択|動作|
|:-----|:-----|:-----|
|Matrix|テキスト|バインドを作成できません。|
|マトリックス|テーブル全体|型 `matrix` のバインドが作成されます。ドキュメントが変更され、コンテンツ コントロールがテーブルをラップする必要があります。 |
|Matrix|テーブル内の範囲|バインドを作成できません。|
|マトリックス|無効な選択項目 (たとえば、複数のオブジェクト、無効なオブジェクトなど)。|バインドを作成できません。|
|テーブル|テキスト|バインドを作成できません。|
|テーブル|テーブル全体|型 `text` のバインドが作成されます。|
|Table|テーブル内の範囲|バインドを作成できません。|
|テーブル|無効な選択項目 (たとえば、複数のオブジェクト、無効なオブジェクトなど)。|バインドを作成できません。|
|テキスト|テーブル全体|型 `text` のバインドが作成されます。|
|テキスト|テーブル内の範囲|バインドを作成できません。|
|テキスト|複数の選択項目|最後の選択項目がコンテンツ コントロール内でラップされ、そのコントロールにバインドされます。 型 `text` のコンテンツ コントロールが作成されます。|
|テキスト|無効な選択項目 (たとえば、複数のオブジェクト、無効なオブジェクトなど)。|バインドを作成できません。|

## <a name="see-also"></a>関連項目

- [Office アドインの開発ライフ サイクル](../overview/office-add-ins.md)
- [Office JavaScript API について](../develop/understanding-the-javascript-api-for-office.md)
- [アプリケーション固有の JavaScript API でのエラー処理](../testing/application-specific-api-error-handling.md)
- [シングル サインオン (SSO) のエラー メッセージのトラブルシューティング](../develop/troubleshoot-sso-in-office-add-ins.md)
- [Office アドインでの開発エラーのトラブルシューティング](../testing/troubleshoot-development-errors.md)
