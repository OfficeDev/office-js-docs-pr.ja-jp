---
title: Office 2013 でのコンテンツ アドインと作業ウィンドウ アドインの Office JavaScript API のサポート
description: Office JavaScript API を使用して、Office 2013 で作業ウィンドウを作成します。
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 42f878d8276bc34f14a69480760aa225dcca6ddb
ms.sourcegitcommit: d8fbe472b35c758753e5d2e4b905a5973e4f7b52
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2022
ms.locfileid: "66229660"
---
# <a name="office-javascript-api-support-for-content-and-task-pane-add-ins-in-office-2013"></a>Office 2013 でのコンテンツ アドインと作業ウィンドウ アドインの Office JavaScript API のサポート

[!include[information about the common API](../includes/alert-common-api-info.md)]

[Office JavaScript API](../reference/javascript-api-for-office.md) を使用して、Office 2013 クライアント アプリケーションの作業ウィンドウまたはコンテンツ アドインを作成できます。 コンテンツと作業ウィンドウのアドインをサポートするオブジェクトとメソッドは、次のように分類されます。

1. **他のOffice アドインと共有される共通オブジェクト。** これらのオブジェクトには [、Office](/javascript/api/office)、[コンテキスト](/javascript/api/office/office.context)、[および AsyncResult が含まれます](/javascript/api/office/office.asyncresult)。 `Office`オブジェクトは、Office JavaScript API のルート オブジェクトです。 オブジェクトは `Context` 、アドインのランタイム環境を表します。 両方`Office`とも`Context`、任意のOffice アドインの基本的なオブジェクトです。 オブジェクトは `AsyncResult` 、ドキュメントでユーザーが選択したものを読み取るメソッドに `getSelectedDataAsync` 返されるデータなど、非同期操作の結果を表します。

2. **Document オブジェクト。** コンテンツ アドインと作業ウィンドウ アドインで使用可能な API の大部分は、[Document](/javascript/api/office/office.document) オブジェクトのメソッド、プロパティ、およびイベントを通して公開されます。 コンテンツまたは作業ウィンドウ アドインは [、Office.context.document](/javascript/api/office/office.context#office-office-context-document-member) プロパティを使用して **Document** オブジェクトにアクセスでき、それを使用して、[Bindings](/javascript/api/office/office.bindings) オブジェクトや [CustomXmlParts](/javascript/api/office/office.customxmlparts) オブジェクト、[getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1))、[setSelectedDataAsync](/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1))、[getFileAsync](/javascript/api/office/office.document#office-office-document-getfileasync-member(1)) メソッドなどのドキュメント内のデータを操作するための API のキー メンバーにアクセスできます。 また、このオブジェクトは`Document`、ドキュメントが読み取り専用か編集モードかを判断するための [mode](/javascript/api/office/office.document#office-office-document-mode-member) プロパティ、現在のドキュメントの URL を取得する [url](/javascript/api/office/office.document#office-office-document-url-member) プロパティ、[および設定](/javascript/api/office/office.settings) オブジェクトへのアクセスを提供します。 このオブジェクトでは `Document` [SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) イベントのイベント ハンドラーの追加もサポートされているため、ユーザーがドキュメント内で選択内容をいつ変更したかを検出できます。

   コンテンツまたは作業ウィンドウ アドインは、DOM とランタイム環境が読み込まれた後にのみオブジェクトにアクセス`Document`できます。通常は[、Office.initialize](/javascript/api/office) イベントのイベント ハンドラーで行われます。 アドインが初期化されるときのイベント フローと、DOM とラインタイムが正常に読み込まれたかどうかの確認方法については、「[DOM とランタイム環境の読み込み](loading-the-dom-and-runtime-environment.md)」を参照してください。

3. **特定の機能を操作するためのオブジェクト。** API の特定の機能を操作するには、次のオブジェクトとメソッドを使用します。

    - [Bindings](/javascript/api/office/office.bindings) オブジェクトのメソッドを使用して、バインドを作成または取得します。また、[Binding](/javascript/api/office/office.binding) オブジェクトのメソッドとプロパティを使用して、データを操作します。

    - [CustomXmlParts](/javascript/api/office/office.customxmlparts)、[CustomXmlPart](/javascript/api/office/office.customxmlpart)、および関連するオブジェクトを使用して、Word 文書内のカスタム XML パーツを作成および操作します。

    - [File](/javascript/api/office/office.file) オブジェクトと [Slice](/javascript/api/office/office.slice) オブジェクトを使用して、文書全体のコピーを作成し、それをチャンクまたは「スライス」に分割してから、それらのスライスに含まれるデータを読み取りまたは転送します。

    - [Settings](/javascript/api/office/office.settings) オブジェクトを使用して、ユーザー設定やアドインの状態などのカスタム データを保存します。

> [!IMPORTANT]
> API メンバーの一部は、コンテンツ アドインと作業ウィンドウ アドインをホスト可能なすべての Office アプリケーションでサポートされているわけではありません。サポートされているメンバーを特定するには、次のいずれかを参照してください。

Office クライアント アプリケーション間での javaScript API のサポートOfficeの概要については、「[Office JavaScript API](understanding-the-javascript-api-for-office.md) について」を参照してください。

## <a name="read-and-write-to-an-active-selection-in-a-document-spreadsheet-or-presentation"></a>文書、スプレッドシート、またはプレゼンテーション内のアクティブな選択範囲の読み取りと書き込み

文書、スプレッドシート、またはプレゼンテーション内のユーザーの現在の選択範囲に対して読み書きをすることができます。 アドインのOffice アプリケーションに応じて、[Document](/javascript/api/office/office.document) オブジェクトの [getSelectedDataAsync メソッドと setSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) メソッドで、パラメーターとして読み取りまたは書き込みを[](/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1))行うデータ構造の種類を指定できます。 たとえば、Word には任意のデータ タイプ (テキスト、HTML、表形式データ、または Office Open XML)、Excel にはテキストと表形式データ、および PowerPoint と Project にはテキストを指定できます。 ユーザーの選択範囲に対する変更を検出するためのイベント ハンドラーを作成することもできます。 次の例では、メソッドを使用して選択範囲からテキストとしてデータを `getSelectedDataAsync` 取得します。


```js
Office.context.document.getSelectedDataAsync(
    Office.CoercionType.Text, function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            write('Action failed. Error: ' + asyncResult.error.message);
        }
        else {
            write('Selected data: ' + asyncResult.value);
        }
    });

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}

```

詳細と例については、「[ドキュメントやスプレッドシート内のアクティブな選択範囲へのデータの読み取りおよび書き込み](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)」を参照してください。

## <a name="bind-to-a-region-in-a-document-or-spreadsheet"></a>ドキュメントまたはスプレッドシート内の領域にバインドする

およびメソッドを`getSelectedDataAsync`使用して、ドキュメント、スプレッドシート、またはプレゼンテーションでユーザーの *現在* の選択範囲を読み書きできます。`setSelectedDataAsync` ただし、ユーザーに選択を要求せずにアドインの複数の実行セッションに渡って文書内の同じ領域にアクセスする場合は、最初にその領域をバインドする必要があります。 そのバインドした領域に対するデータおよび選択範囲変更イベントにサブスクライブすることもできます。

バインドは、[Bindings](/javascript/api/office/office.bindings#office-office-bindings-addfromnameditemasync-member(1)) オブジェクトの [addFromNamedItemAsync](/javascript/api/office/office.bindings#office-office-bindings-addfrompromptasync-member(1)) メソッド、[addFromPromptAsync](/javascript/api/office/office.bindings#office-office-bindings-addfromselectionasync-member(1)) メソッド、または [addFromSelectionAsync](/javascript/api/office/office.bindings) メソッドを使用して追加できます。これらのメソッドは、バインド内のデータにアクセスするため、あるいは、データ変更または選択範囲変更イベントにサブスクライブするために使用可能な識別子を返します。

次に、メソッドを使用して、ドキュメント内で現在選択されているテキストにバインドを追加する例を `Bindings.addFromSelectionAsync` 示します。

```js
Office.context.document.bindings.addFromSelectionAsync(
    Office.BindingType.Text, { id: 'myBinding' }, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write('Added new binding with type: ' +
            asyncResult.value.type + ' and id: ' + asyncResult.value.id);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

詳細と例については、「[ドキュメントまたはスプレッドシート内の領域へのバインド](bind-to-regions-in-a-document-or-spreadsheet.md)」を参照してください。

## <a name="get-entire-documents"></a>ドキュメント全体を取得する

作業ウィンドウ アドインが PowerPoint または Word で実行される場合は、[Document.getFileAsync](/javascript/api/office/office.document#office-office-document-getfileasync-member(1)) メソッド、[File.getSliceAsync](/javascript/api/office/office.file#office-office-file-getsliceasync-member(1)) メソッド、および [File.closeAsync](/javascript/api/office/office.file#office-office-file-closeasync-member(1)) メソッドを使用して、プレゼンテーションまたは文書全体を取得できます。

呼び出 `Document.getFileAsync` すと、 [File](/javascript/api/office/office.file) オブジェクト内のドキュメントのコピーが取得されます。 オブジェクトは `File` 、Slice オブジェクトとして表される "チャンク" でドキュメントへのアクセス [を](/javascript/api/office/office.slice) 提供します。 呼び出すとき`getFileAsync`は、ファイルの種類 (テキストまたは圧縮された Open Office XML 形式) とスライスのサイズ (最大 4 MB) を指定できます。 オブジェクトの内容に`File`アクセスするには、[Slice.data](/javascript/api/office/office.slice#office-office-slice-data-member) プロパティの生データを返すオブジェクトを呼び出`File.getSliceAsync`します。 圧縮形式を指定した場合は、ファイル データがバイト配列で返されます。 ファイルを Web サービスに転送する場合は、圧縮生データを base64 エンコード文字列に変換してから送信できます。 最後に、ファイルのスライスの取得が完了したら、メソッドを `File.closeAsync` 使用してドキュメントを閉じます。

詳細については、[PowerPoint や Word 用のアドインからドキュメント全体を取得する](../word/get-the-whole-document-from-an-add-in-for-word.md)方法を参照してください。

## <a name="read-and-write-custom-xml-parts-of-a-word-document"></a>Word 文書のカスタム XML 部分の読み取りと書き込み

Open Office XML ファイル形式とコンテンツ コントロールを使用すれば、Word 文書にカスタム XML パーツを追加して、その文書内のコンテンツ コントロールに XML パーツ内の要素をバインドすることができます。文書を開くと、Word がバインドされたコンテンツ コントロールを読み取り、カスタム XML パーツからのデータを自動的に設定します。ユーザーは、コンテンツ コントロールにデータを書き込むこともできます。ユーザーが文書を保存すると、コントロール内のデータがバインドされた XML パーツに保存されます。Word 用の作業ウィンドウ アドインは、[Document.customXmlParts](/javascript/api/office/office.document#office-office-document-customxmlparts-member) プロパティ、[CustomXmlParts](/javascript/api/office/office.customxmlparts) オブジェクト、[CustomXmlPart](/javascript/api/office/office.customxmlpart) オブジェクト、および [CustomXmlNode](/javascript/api/office/office.customxmlnode) オブジェクトを使用して、文書に対して動的にデータを読み書きすることができます。

カスタム XML パーツは名前空間に関連付けることができます。名前空間内のカスタム XML パーツからデータを取得するには、[CustomXmlParts.getByNamespaceAsync](/javascript/api/office/office.customxmlparts#office-office-customxmlparts-getbynamespaceasync-member(1)) メソッドを使用します。

[CustomXmlParts.getByIdAsync](/javascript/api/office/office.customxmlparts#office-office-customxmlparts-getbyidasync-member(1)) メソッドを使用して、GUID でカスタム XML パーツにアクセスすることもできます。カスタム XML パーツを取得したら、[CustomXmlPart.getXmlAsync](/javascript/api/office/office.customxmlpart#office-office-customxmlpart-getxmlasync-member(1)) メソッドを使用して XML データを取得します。

ドキュメントに新しいカスタム XML パーツを追加するには、このプロパティを `Document.customXmlParts` 使用してドキュメント内のカスタム XML パーツを取得し、 [CustomXmlParts.addAsync](/javascript/api/office/office.customxmlparts#office-office-customxmlparts-addasync-member(1)) メソッドを呼び出します。

作業ウィンドウ アドインを使用してカスタム XML パーツを管理する方法の詳細については、「Word アドインで xml を[開くOffice使用するタイミングと方法について理解する](../word/create-better-add-ins-for-word-with-office-open-xml.md)」を参照してください。

## <a name="persisting-add-in-settings"></a>アドイン設定を保存する

多くの場合、ユーザー設定やアドインの状態など、アドインのカスタム データを保存し、次回、アドインを開いたとき、そのデータにアクセスする必要があります。 一般的な Web プログラミング手法を利用し、ブラウザーの Cookie や HTML 5 Web ストレージなど、そのデータを保存できます。 あるいは、アドインを Excel、PowerPoint、Word で実行する場合、[Settings](/javascript/api/office/office.settings) オブジェクトのメソッドを使用できます。 オブジェクトを使用して作成された `Settings` データは、アドインが挿入および保存されたスプレッドシート、プレゼンテーション、またはドキュメントに格納されます。 このデータは、それを作成したアドインでのみ利用できます。

ドキュメントが格納されているサーバーへのラウンドトリップを回避するために、オブジェクトで作成されたデータは実行時に `Settings` メモリ内で管理されます。 過去に保存した設定データがアドインの初期化時にメモリに読み込まれ、そのデータに対する変更は [Settings.saveAsync](/javascript/api/office/office.settings#office-office-settings-saveasync-member(1)) メソッドを呼び出したときにのみ文書に保存されます。 内部的に、データはシリアル化された JSON オブジェクト内に名前と値のペアとして保存されます。 データのメモリ内コピーに対してアイテムの読み取り、書き込み、および削除を実行するには、[Settings](/javascript/api/office/office.settings#office-office-settings-get-member(1)) オブジェクトの [get](/javascript/api/office/office.settings#office-office-settings-set-member(1)) メソッド、[set](/javascript/api/office/office.settings#office-office-settings-remove-member(1)) メソッド、および **remove** メソッドを使用します。 次のコード行は、`themeColor` という名前の設定を作成して、その値を 'green' に設定する方法を示しています。

```js
Office.context.document.settings.set('themeColor', 'green');
```

メソッドと`remove`共`set`に作成または削除された設定データは、データのメモリ内コピーに対して動作するため、アドインが操作しているドキュメントに設定データの変更を保持するために呼び出す`saveAsync`必要があります。

オブジェクトのメソッドを使用したカスタム データの操作の詳細については、「アドインの`Settings`[状態と設定の永続化](persisting-add-in-state-and-settings.md)」を参照してください。

## <a name="read-properties-of-a-project-document"></a>プロジェクト ドキュメントのプロパティの読み取り

作業ウィンドウ アドインが Project で動作する場合は、そのアドインでアクティブ プロジェクト内のプロジェクト フィールド、リソース、およびタスク フィールドの一部からデータを読み取ることができます。 これを行うには、[ProjectDocument](/javascript/api/office/office.document) オブジェクトのメソッドとイベントを使用します。これにより、オブジェクトが`Document`拡張され、追加のProject固有の機能が提供されます。

Project のデータの読み取り操作の例については、「[テキスト エディターを使用して Project 2013 用の作業ウィンドウ アドインを初めて作成する](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)」を参照してください。

## <a name="permissions-model-and-governance"></a>アクセス許可モデルとガバナンス

アドインでは、マニフェスト内の要素を`Permissions`使用して、Office JavaScript API から必要な機能レベルへのアクセス許可を要求します。 たとえば、アドインでドキュメントへの読み取り/書き込みアクセスが必要な場合、そのマニフェストは要素の`Permissions`テキスト値として指定`ReadWriteDocument`する必要があります。 アクセス許可はユーザーのプライバシーとセキュリティを保護するために存在しているので、ベスト プラクティスとしては、その機能に必要な最低限のアクセス許可を要求することをお勧めします。 次の例は、作業ウィンドウのマニフェストで **ReadDocument** アクセス許可を要求する方法を示しています。

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
 xsi:type="TaskPaneApp">
???<!-- Other manifest elements omitted. -->
  <Permissions>ReadDocument</Permissions>
???
</OfficeApp>

```

詳細については、「 [アドインでの API の使用に対するアクセス許可の要求](requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)」を参照してください。

## <a name="see-also"></a>関連項目

- [Office の JavaScript API](../reference/javascript-api-for-office.md)
- [Office アドイン マニフェストのスキーマ参照](../develop/add-in-manifests.md)
- [Office アドインでのユーザー エラーのトラブルシューティング](../testing/testing-and-troubleshooting.md)
