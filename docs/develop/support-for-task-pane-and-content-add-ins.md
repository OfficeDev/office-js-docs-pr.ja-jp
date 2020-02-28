---
title: Office 2013 でのコンテンツ アドインと作業ウィンドウ アドインの Office JavaScript API のサポート
description: ''
ms.date: 02/27/2020
localization_priority: Normal
ms.openlocfilehash: a9eb67ca78f89888860ff3ed11ae1632ff62b690
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/27/2020
ms.locfileid: "42323823"
---
# <a name="office-javascript-api-support-for-content-and-task-pane-add-ins-in-office-2013"></a>Office 2013 でのコンテンツ アドインと作業ウィンドウ アドインの Office JavaScript API のサポート

[!include[information about the common API](../includes/alert-common-api-info.md)]

[Office JavaScript API](../reference/javascript-api-for-office.md) を使用して、Office 2013 ホスト アプリケーションの作業ウィンドウやコンテンツ用のアドインを作成できます。コンテンツと作業ウィンドウのアドインをサポートするオブジェクトとメソッドは、次のように分類されます。

1. **他の Office アドインと共有されている共通オブジェクト。** これらのオブジェクトには、 [Office](/javascript/api/office)、 [Context](/javascript/api/office/office.context)、および[AsyncResult](/javascript/api/office/office.asyncresult)があります。 `Office`オブジェクトは、OFFICE JavaScript API のルートオブジェクトです。 オブジェクト`Context`は、アドインのランタイム環境を表します。 `Office`および`Context`は、Office アドインの基本的なオブジェクトです。 オブジェクト`AsyncResult`は、 `getSelectedDataAsync`メソッドに返されるデータなど、非同期操作の結果を表します。これは、ユーザーがドキュメント内で選択したものを読み取ります。

2. **Document オブジェクト。** コンテンツ アドインと作業ウィンドウ アドインで使用可能な API の大部分は、[Document](/javascript/api/office/office.document) オブジェクトのメソッド、プロパティ、およびイベントを通して公開されます。 コンテンツアドインまたは作業ウィンドウアドインは、 [Office](/javascript/api/office/office.context#document)のプロパティを使用して**document**オブジェクトにアクセスできます。また、このプロパティを使用して、ドキュメント内のデータを操作するための API のキーメンバー ( [Bindings](/javascript/api/office/office.bindings)オブジェクト、 [Customxmlparts](/javascript/api/office/office.customxmlparts)オブジェクト、 [Getselecteddataasync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-)、 [setselecteddataasync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-)、および[getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-)メソッドなど) にアクセスできます。 オブジェクト`Document`には、ドキュメントが読み取り専用であるか、編集モードであるかを決定する[mode](/javascript/api/office/office.document#mode)プロパティ、現在のドキュメントの url を取得するための[url](/javascript/api/office/office.document#url)プロパティ、および[Settings](/javascript/api/office/office.settings)オブジェクトへのアクセスも用意されています。 `Document`オブジェクトでは、 [selectionchanged](/javascript/api/office/office.documentselectionchangedeventargs)イベントのイベントハンドラーの追加もサポートされているため、ユーザーが文書内の選択範囲を変更したときに検出できます。

   コンテンツアドインまたは作業ウィンドウアドインは、DOM と`Document`ランタイム環境が読み込まれた後にのみ、オブジェクトにアクセスできます。通常、このイベントハンドラーは、 [Office の initialize](/javascript/api/office)イベントのイベントハンドラーです。 アドインが初期化されるときのイベント フローと、DOM とラインタイムが正常に読み込まれたかどうかの確認方法については、「[DOM とランタイム環境の読み込み](loading-the-dom-and-runtime-environment.md)」を参照してください。

3. **特定の機能を操作するためのオブジェクト。** API の特定の機能を操作するには、次のオブジェクトとメソッドを使用します。

    - [Bindings](/javascript/api/office/office.bindings) オブジェクトのメソッドを使用して、バインドを作成または取得します。また、[Binding](/javascript/api/office/office.binding) オブジェクトのメソッドとプロパティを使用して、データを操作します。

    - [CustomXmlParts](/javascript/api/office/office.customxmlparts)、[CustomXmlPart](/javascript/api/office/office.customxmlpart)、および関連するオブジェクトを使用して、Word 文書内のカスタム XML パーツを作成および操作します。

    - [File](/javascript/api/office/office.file) オブジェクトと [Slice](/javascript/api/office/office.slice) オブジェクトを使用して、文書全体のコピーを作成し、それをチャンクまたは「スライス」に分割してから、それらのスライスに含まれるデータを読み取りまたは転送します。

    - [Settings](/javascript/api/office/office.settings) オブジェクトを使用して、ユーザー設定やアドインの状態などのカスタム データを保存します。


> [!IMPORTANT]
> API メンバーの一部は、コンテンツ アドインと作業ウィンドウ アドインをホスト可能なすべての Office アプリケーションでサポートされているわけではありません。サポートされているメンバーを特定するには、次のいずれかを参照してください。

Office ホストアプリケーション全体にわたる Office JavaScript API サポートの概要については、「 [Office JAVASCRIPT api につい](understanding-the-javascript-api-for-office.md)て」を参照してください。


## <a name="reading-and-writing-to-an-active-selection"></a>アクティブな選択範囲の読み取りと書き込み

文書、スプレッドシート、またはプレゼンテーション内のユーザーの現在の選択範囲に対して読み書きをすることができます。 アドイン用のホスト アプリケーションに応じて、[Document](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) オブジェクトの [getSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) メソッドと [setSelectedDataAsync](/javascript/api/office/office.document) メソッド内のパラメーターとして読み書きするデータ構造のタイプを指定できます。 たとえば、Word には任意のデータ タイプ (テキスト、HTML、表形式データ、または Office Open XML)、Excel にはテキストと表形式データ、および PowerPoint と Project にはテキストを指定できます。 ユーザーの選択範囲に対する変更を検出するためのイベント ハンドラーを作成することもできます。 次の例では、 `getSelectedDataAsync`メソッドを使用して選択範囲からデータをテキストとして取得します。


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

詳細と例については、「[文書またはスプレッドシート内のアクティブな選択範囲へのデータの読み取りと書き込み](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)」を参照してください。


## <a name="binding-to-a-region-in-a-document-or-spreadsheet"></a>文書またはスプレッドシート内の領域へのバインド

`getSelectedDataAsync`およびメソッドを使用する`setSelectedDataAsync`と、ドキュメント、スプレッドシート、またはプレゼンテーションで、ユーザーの*現在*の選択範囲の読み取りや書き込みを行うことができます。 ただし、ユーザーに選択を要求せずにアドインの複数の実行セッションに渡って文書内の同じ領域にアクセスする場合は、最初にその領域をバインドする必要があります。 そのバインドした領域に対するデータおよび選択範囲変更イベントにサブスクライブすることもできます。

バインドは、[Bindings](/javascript/api/office/office.bindings#addfromnameditemasync-itemname--bindingtype--options--callback-) オブジェクトの [addFromNamedItemAsync](/javascript/api/office/office.bindings#addfrompromptasync-bindingtype--options--callback-) メソッド、[addFromPromptAsync](/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-) メソッド、または [addFromSelectionAsync](/javascript/api/office/office.bindings) メソッドを使用して追加できます。これらのメソッドは、バインド内のデータにアクセスするため、あるいは、データ変更または選択範囲変更イベントにサブスクライブするために使用可能な識別子を返します。

次の例は、 `Bindings.addFromSelectionAsync`メソッドを使用して、ドキュメントで現在選択されているテキストにバインドを追加します。



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

詳細と例については、「[文書またはスプレッドシート内の領域へのバインド](bind-to-regions-in-a-document-or-spreadsheet.md)」を参照してください。


## <a name="getting-entire-documents"></a>文書全体の取得

作業ウィンドウ アドインが PowerPoint または Word で実行される場合は、[Document.getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-) メソッド、[File.getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-) メソッド、および [File.closeAsync](/javascript/api/office/office.file#closeasync-callback-) メソッドを使用して、プレゼンテーションまたは文書全体を取得できます。

を呼び出す`Document.getFileAsync`と、[ファイル](/javascript/api/office/office.file)オブジェクトにドキュメントのコピーが取得されます。 オブジェクト`File`は、 [Slice](/javascript/api/office/office.slice)オブジェクトとして表される "チャンク" 内のドキュメントへのアクセスを提供します。 を呼び出す`getFileAsync`ときに、ファイルの種類 (テキストまたは圧縮された Office XML 形式) と、スライスのサイズ (最大 4mb) を指定できます。 `File`オブジェクトの内容にアクセスするために、を呼び`File.getSliceAsync`出すと、 [data](/javascript/api/office/office.slice#data)プロパティにある生のデータが返されます。 圧縮形式を指定した場合は、ファイル データがバイト配列で返されます。 ファイルを Web サービスに転送する場合は、圧縮生データを base64 エンコード文字列に変換してから送信できます。 最後に、ファイルのスライスの取得が終了したら、 `File.closeAsync`メソッドを使用してドキュメントを閉じます。

詳細については、「[PowerPoint または Word 用アドインからドキュメント全体を取得する方法](../word/get-the-whole-document-from-an-add-in-for-word.md)」を参照してください。


## <a name="reading-and-writing-custom-xml-parts-of-a-word-document"></a>Word 文書のカスタム XML パーツの読み取りと書き込み

Open Office XML ファイル形式とコンテンツ コントロールを使用すれば、Word 文書にカスタム XML パーツを追加して、その文書内のコンテンツ コントロールに XML パーツ内の要素をバインドすることができます。文書を開くと、Word がバインドされたコンテンツ コントロールを読み取り、カスタム XML パーツからのデータを自動的に設定します。ユーザーは、コンテンツ コントロールにデータを書き込むこともできます。ユーザーが文書を保存すると、コントロール内のデータがバインドされた XML パーツに保存されます。Word 用の作業ウィンドウ アドインは、[Document.customXmlParts](/javascript/api/office/office.document#customxmlparts) プロパティ、[CustomXmlParts](/javascript/api/office/office.customxmlparts) オブジェクト、[CustomXmlPart](/javascript/api/office/office.customxmlpart) オブジェクト、および [CustomXmlNode](/javascript/api/office/office.customxmlnode) オブジェクトを使用して、文書に対して動的にデータを読み書きすることができます。

カスタム XML パーツは名前空間に関連付けることができます。名前空間内のカスタム XML パーツからデータを取得するには、[CustomXmlParts.getByNamespaceAsync](/javascript/api/office/office.customxmlparts#getbynamespaceasync-ns--options--callback-) メソッドを使用します。

[CustomXmlParts.getByIdAsync](/javascript/api/office/office.customxmlparts#getbyidasync-id--options--callback-) メソッドを使用して、GUID でカスタム XML パーツにアクセスすることもできます。カスタム XML パーツを取得したら、[CustomXmlPart.getXmlAsync](/javascript/api/office/office.customxmlpart#getxmlasync-options--callback-) メソッドを使用して XML データを取得します。

新しいカスタム XML 部分をドキュメントに追加するには、 `Document.customXmlParts`プロパティを使用してドキュメント内のカスタム xml 部分を取得し、 [customxmlparts](/javascript/api/office/office.customxmlparts#addasync-xml--options--callback-)メソッドを呼び出します。

作業ウィンドウ アドインでのカスタム XML パーツの操作方法の詳細については、「[Office Open XML を使用してより良い Word 用アドインを作成する](../word/create-better-add-ins-for-word-with-office-open-xml.md)」を参照してください。


## <a name="persisting-add-in-settings"></a>アドイン設定を保存する


多くの場合、ユーザー設定やアドインの状態など、アドインのカスタム データを保存し、次回、アドインを開いたとき、そのデータにアクセスする必要があります。 一般的な Web プログラミング手法を利用し、ブラウザーの Cookie や HTML 5 Web ストレージなど、そのデータを保存できます。 あるいは、アドインを Excel、PowerPoint、Word で実行する場合、[Settings](/javascript/api/office/office.settings) オブジェクトのメソッドを使用できます。 `Settings`オブジェクトを使用して作成されたデータは、アドインが挿入されてに保存されたスプレッドシート、プレゼンテーション、またはドキュメントに格納されます。 このデータは、それを作成したアドインでのみ利用できます。

ドキュメントが格納されているサーバーへのラウンドトリップを回避するため`Settings`に、オブジェクトを使用して作成されたデータは実行時にメモリで管理されます。 過去に保存した設定データがアドインの初期化時にメモリに読み込まれ、そのデータに対する変更は [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) メソッドを呼び出したときにのみ文書に保存されます。 内部的に、データはシリアル化された JSON オブジェクト内に名前と値のペアとして保存されます。 データのメモリ内コピーに対してアイテムの読み取り、書き込み、および削除を実行するには、[Settings](/javascript/api/office/office.settings#get-name-) オブジェクトの [get](/javascript/api/office/office.settings#set-name--value-) メソッド、[set](/javascript/api/office/office.settings#remove-name-) メソッド、および **remove** メソッドを使用します。 次のコード行は、`themeColor` という名前の設定を作成して、その値を 'green' に設定する方法を示しています。




```js
Office.context.document.settings.set('themeColor', 'green');
```

およびメソッドを使用して作成または削除された設定データは、データのメモリ内コピーに対して`saveAsync`動作するので、設定データへの変更をアドインが作業しているドキュメントに保持するために、を呼び出す必要があります。 `remove` `set`

オブジェクトのメソッドを使用したカスタムデータの使用の詳細については、「[アドインの状態と設定を保持](persisting-add-in-state-and-settings.md)する」を参照してください。 `Settings`


## <a name="reading-properties-of-a-project-document"></a>プロジェクト文書のプロパティの読み取り

作業ウィンドウ アドインが Project で動作する場合は、そのアドインでアクティブ プロジェクト内のプロジェクト フィールド、リソース、およびタスク フィールドの一部からデータを読み取ることができます。 そのためには、 `Document`オブジェクトを拡張する[projectdocument](/javascript/api/office/office.document)オブジェクトのメソッドとイベントを使用します。これにより、追加のプロジェクト固有の機能が提供されます。

Project のデータの読み取り操作の例については、「[テキスト エディターを使用して Project 2013 用の作業ウィンドウ アドインを初めて作成する](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)」を参照してください。


## <a name="permissions-model-and-governance"></a>アクセス許可モデルとガバナンス

アドインでは、マニフェスト内`Permissions`の要素を使用して、OFFICE JavaScript API から必要な機能のレベルにアクセスするためのアクセス許可を要求します。 たとえば、アドインでドキュメントに対する読み取り/書き込みアクセス権が必要な場合、そのマニフェストは、 `ReadWriteDocument` `Permissions`要素のテキスト値として指定する必要があります。 アクセス許可はユーザーのプライバシーとセキュリティを保護するために存在しているので、ベスト プラクティスとしては、その機能に必要な最低限のアクセス許可を要求することをお勧めします。 次の例は、作業ウィンドウのマニフェストで **ReadDocument** アクセス許可を要求する方法を示しています。


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

詳細については、「[アドインで API を使用するためのアクセス許可を要求](requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)する」を参照してください。


## <a name="see-also"></a>関連項目

- [Office の JavaScript API](/office/dev/add-ins/reference/javascript-api-for-office)
- [Office アドイン マニフェストのスキーマ参照](../develop/add-in-manifests.md)
- [Office アドインでのユーザー エラーのトラブルシューティング](../testing/testing-and-troubleshooting.md)
