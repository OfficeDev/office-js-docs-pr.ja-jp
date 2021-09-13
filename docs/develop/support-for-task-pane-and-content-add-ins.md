---
title: Office 2013 でのコンテンツ アドインと作業ウィンドウ アドインの Office JavaScript API のサポート
description: 2013 Office JavaScript API を使用して作業ウィンドウを作成Officeします。
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 8af93e7cd0ba527c72a4e6e721e30fb9739dda6a
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59149941"
---
# <a name="office-javascript-api-support-for-content-and-task-pane-add-ins-in-office-2013"></a>Office 2013 でのコンテンツ アドインと作業ウィンドウ アドインの Office JavaScript API のサポート

[!include[information about the common API](../includes/alert-common-api-info.md)]

[JavaScript API](../reference/javascript-api-for-office.md)の Officeを使用して、2013 クライアント アプリケーション用の作業ウィンドウまたはコンテンツ アドインOffice作成できます。 コンテンツと作業ウィンドウのアドインをサポートするオブジェクトとメソッドは、次のように分類されます。

1. **他のアドインと共有Office共通のオブジェクト。** これらのオブジェクトには [、Office、Context、](/javascript/api/office)[](/javascript/api/office/office.context)および [AsyncResult が含まれます](/javascript/api/office/office.asyncresult)。 オブジェクト `Office` は、JavaScript API のOfficeオブジェクトです。 オブジェクト `Context` は、アドインのランタイム環境を表します。 両方 `Office` とも `Context` 、任意のアドインの基本的Officeオブジェクトです。 オブジェクトは、メソッドに返されるデータなどの非同期操作の結果を表し、ユーザーがドキュメントで選択した値 `AsyncResult` `getSelectedDataAsync` を読み取ります。

2. **Document オブジェクト。** コンテンツ アドインと作業ウィンドウ アドインで使用可能な API の大部分は、[Document](/javascript/api/office/office.document) オブジェクトのメソッド、プロパティ、およびイベントを通して公開されます。 コンテンツ アドインまたは作業ウィンドウ アドインは [、Office.context.document](/javascript/api/office/office.context#document)プロパティを使用して **Document** オブジェクトにアクセスし、それを介して [、Bindings](/javascript/api/office/office.bindings)オブジェクトや [CustomXmlParts](/javascript/api/office/office.customxmlparts)オブジェクト [、getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__options__callback_)メソッド [、getFileAsync](/javascript/api/office/office.document#setSelectedDataAsync_data__options__callback_)メソッドなどのドキュメント内のデータを操作するための API の主要メンバーに [](/javascript/api/office/office.document#getFileAsync_fileType__options__callback_)アクセスできます。 オブジェクトには、ドキュメントが読み取り専用か編集モードかを判断するための mode プロパティ、現在のドキュメントの URL を取得するための url プロパティ、および 設定 オブジェクトへのアクセスも `Document` [提供](/javascript/api/office/office.settings)されます。 [](/javascript/api/office/office.document#mode) [](/javascript/api/office/office.document#url) オブジェクトは SelectionChanged イベントのイベント ハンドラーの追加もサポートしています。そのため、ユーザーがドキュメント内で選択内容を変更した場合 `Document` を検出できます。 [](/javascript/api/office/office.documentselectionchangedeventargs)

   コンテンツアドインまたは作業ウィンドウ アドインは、DOM およびランタイム環境が読み込まれた後にのみオブジェクトにアクセスできます(通常は `Document` [Office.initialize イベントのイベント ハンドラー](/javascript/api/office)で)。 アドインが初期化されるときのイベント フローと、DOM とラインタイムが正常に読み込まれたかどうかの確認方法については、「[DOM とランタイム環境の読み込み](loading-the-dom-and-runtime-environment.md)」を参照してください。

3. **特定の機能を操作するためのオブジェクト。** API の特定の機能を操作するには、次のオブジェクトとメソッドを使用します。

    - [Bindings](/javascript/api/office/office.bindings) オブジェクトのメソッドを使用して、バインドを作成または取得します。また、[Binding](/javascript/api/office/office.binding) オブジェクトのメソッドとプロパティを使用して、データを操作します。

    - [CustomXmlParts](/javascript/api/office/office.customxmlparts)、[CustomXmlPart](/javascript/api/office/office.customxmlpart)、および関連するオブジェクトを使用して、Word 文書内のカスタム XML パーツを作成および操作します。

    - [File](/javascript/api/office/office.file) オブジェクトと [Slice](/javascript/api/office/office.slice) オブジェクトを使用して、文書全体のコピーを作成し、それをチャンクまたは「スライス」に分割してから、それらのスライスに含まれるデータを読み取りまたは転送します。

    - [Settings](/javascript/api/office/office.settings) オブジェクトを使用して、ユーザー設定やアドインの状態などのカスタム データを保存します。

> [!IMPORTANT]
> API メンバーの一部は、コンテンツ アドインと作業ウィンドウ アドインをホスト可能なすべての Office アプリケーションでサポートされているわけではありません。サポートされているメンバーを特定するには、次のいずれかを参照してください。

クライアント アプリケーション全体での JavaScript API Officeの概要Office [JavaScript API](understanding-the-javascript-api-for-office.md)についてを参照Office参照してください。

## <a name="read-and-write-to-an-active-selection-in-a-document-spreadsheet-or-presentation"></a>ドキュメント、スプレッドシート、またはプレゼンテーションのアクティブな選択範囲に対する読み取りおよび書き込み

文書、スプレッドシート、またはプレゼンテーション内のユーザーの現在の選択範囲に対して読み書きをすることができます。 アドインの Office アプリケーションに応じて[、Document](/javascript/api/office/office.document)オブジェクトの[getSelectedDataAsync メソッドおよび setSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__options__callback_)メソッドで、パラメーターとして読み[](/javascript/api/office/office.document#setSelectedDataAsync_data__options__callback_)取りまたは書き込みを行うデータ構造の種類を指定できます。 たとえば、Word には任意のデータ タイプ (テキスト、HTML、表形式データ、または Office Open XML)、Excel にはテキストと表形式データ、および PowerPoint と Project にはテキストを指定できます。 ユーザーの選択範囲に対する変更を検出するためのイベント ハンドラーを作成することもできます。 次の使用例は、メソッドを使用して選択範囲のデータをテキストとして取得 `getSelectedDataAsync` します。


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

and メソッドを使用すると、ドキュメント、スプレッドシート、プレゼンテーションでユーザーの現在の選択内容を読み取りまたは `getSelectedDataAsync` `setSelectedDataAsync` 書き込みできます。  ただし、ユーザーに選択を要求せずにアドインの複数の実行セッションに渡って文書内の同じ領域にアクセスする場合は、最初にその領域をバインドする必要があります。 そのバインドした領域に対するデータおよび選択範囲変更イベントにサブスクライブすることもできます。

バインドは、[Bindings](/javascript/api/office/office.bindings#addFromNamedItemAsync_itemName__bindingType__options__callback_) オブジェクトの [addFromNamedItemAsync](/javascript/api/office/office.bindings#addFromPromptAsync_bindingType__options__callback_) メソッド、[addFromPromptAsync](/javascript/api/office/office.bindings#addFromSelectionAsync_bindingType__options__callback_) メソッド、または [addFromSelectionAsync](/javascript/api/office/office.bindings) メソッドを使用して追加できます。これらのメソッドは、バインド内のデータにアクセスするため、あるいは、データ変更または選択範囲変更イベントにサブスクライブするために使用可能な識別子を返します。

次に、メソッドを使用して、ドキュメント内で現在選択されているテキストにバインドを追加する例を示 `Bindings.addFromSelectionAsync` します。

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

作業ウィンドウ アドインが PowerPoint または Word で実行される場合は、[Document.getFileAsync](/javascript/api/office/office.document#getFileAsync_fileType__options__callback_) メソッド、[File.getSliceAsync](/javascript/api/office/office.file#getSliceAsync_sliceIndex__callback_) メソッド、および [File.closeAsync](/javascript/api/office/office.file#closeAsync_callback_) メソッドを使用して、プレゼンテーションまたは文書全体を取得できます。

呼び出し `Document.getFileAsync` 時に、File オブジェクト内のドキュメントのコピーを [取得](/javascript/api/office/office.file) します。 オブジェクト `File` は、Slice オブジェクトとして表される "チャンク" でドキュメントへの [アクセスを提供](/javascript/api/office/office.slice) します。 呼び出す場合は、ファイルの種類 (テキストまたは圧縮された Open Office XML 形式)、スライスのサイズ `getFileAsync` (最大 4 MB) を指定できます。 オブジェクトの内容にアクセスするには、Slice.data プロパティの生データを返す `File` `File.getSliceAsync` 呼 [び出](/javascript/api/office/office.slice#data) しを行います。 圧縮形式を指定した場合は、ファイル データがバイト配列で返されます。 ファイルを Web サービスに転送する場合は、圧縮生データを base64 エンコード文字列に変換してから送信できます。 最後に、ファイルのスライスの取得が完了したら、メソッドを使用して `File.closeAsync` ドキュメントを閉じます。

詳細については、[PowerPoint や Word 用のアドインからドキュメント全体を取得する](../word/get-the-whole-document-from-an-add-in-for-word.md)方法を参照してください。

## <a name="read-and-write-custom-xml-parts-of-a-word-document"></a>Word ドキュメントのカスタム XML パーツの読み取りおよび書き込み

Open Office XML ファイル形式とコンテンツ コントロールを使用すれば、Word 文書にカスタム XML パーツを追加して、その文書内のコンテンツ コントロールに XML パーツ内の要素をバインドすることができます。文書を開くと、Word がバインドされたコンテンツ コントロールを読み取り、カスタム XML パーツからのデータを自動的に設定します。ユーザーは、コンテンツ コントロールにデータを書き込むこともできます。ユーザーが文書を保存すると、コントロール内のデータがバインドされた XML パーツに保存されます。Word 用の作業ウィンドウ アドインは、[Document.customXmlParts](/javascript/api/office/office.document#customXmlParts) プロパティ、[CustomXmlParts](/javascript/api/office/office.customxmlparts) オブジェクト、[CustomXmlPart](/javascript/api/office/office.customxmlpart) オブジェクト、および [CustomXmlNode](/javascript/api/office/office.customxmlnode) オブジェクトを使用して、文書に対して動的にデータを読み書きすることができます。

カスタム XML パーツは名前空間に関連付けることができます。名前空間内のカスタム XML パーツからデータを取得するには、[CustomXmlParts.getByNamespaceAsync](/javascript/api/office/office.customxmlparts#getByNamespaceAsync_ns__options__callback_) メソッドを使用します。

[CustomXmlParts.getByIdAsync](/javascript/api/office/office.customxmlparts#getByIdAsync_id__options__callback_) メソッドを使用して、GUID でカスタム XML パーツにアクセスすることもできます。カスタム XML パーツを取得したら、[CustomXmlPart.getXmlAsync](/javascript/api/office/office.customxmlpart#getXmlAsync_options__callback_) メソッドを使用して XML データを取得します。

新しいカスタム XML パーツをドキュメントに追加するには、このプロパティを使用して、ドキュメント内のカスタム XML パーツを取得し `Document.customXmlParts` [、CustomXmlParts.addAsync メソッドを呼び出](/javascript/api/office/office.customxmlparts#addAsync_xml__options__callback_) します。

作業ウィンドウ アドインでのカスタム XML パーツの操作方法の詳細については、「[Office Open XML を使用してより良い Word 用アドインを作成する](../word/create-better-add-ins-for-word-with-office-open-xml.md)」を参照してください。

## <a name="persisting-add-in-settings"></a>アドイン設定を保存する

多くの場合、ユーザー設定やアドインの状態など、アドインのカスタム データを保存し、次回、アドインを開いたとき、そのデータにアクセスする必要があります。 一般的な Web プログラミング手法を利用し、ブラウザーの Cookie や HTML 5 Web ストレージなど、そのデータを保存できます。 あるいは、アドインを Excel、PowerPoint、Word で実行する場合、[Settings](/javascript/api/office/office.settings) オブジェクトのメソッドを使用できます。 `Settings` オブジェクトで作成したデータは、アドインを挿入して保存したスプレッドシート、プレゼンテーション、文書に保存されます。 このデータは、それを作成したアドインでのみ利用できます。

ドキュメントが格納されているサーバーへのラウンドトリップを回避するために、オブジェクトで作成されたデータは実行時に `Settings` メモリで管理されます。 過去に保存した設定データがアドインの初期化時にメモリに読み込まれ、そのデータに対する変更は [Settings.saveAsync](/javascript/api/office/office.settings#saveAsync_options__callback_) メソッドを呼び出したときにのみ文書に保存されます。 内部的に、データはシリアル化された JSON オブジェクト内に名前と値のペアとして保存されます。 データのメモリ内コピーに対してアイテムの読み取り、書き込み、および削除を実行するには、[Settings](/javascript/api/office/office.settings#get_name_) オブジェクトの [get](/javascript/api/office/office.settings#set_name__value_) メソッド、[set](/javascript/api/office/office.settings#remove_name_) メソッド、および **remove** メソッドを使用します。 次のコード行は、`themeColor` という名前の設定を作成して、その値を 'green' に設定する方法を示しています。

```js
Office.context.document.settings.set('themeColor', 'green');
```

and メソッドを使用して作成または削除された設定データは、メモリ内のデータ コピーに対して動作しますので、アドインが操作しているドキュメントに設定データの変更を保持するために呼び出す必要があります。 `set` `remove` `saveAsync`

オブジェクトのメソッドを使用したカスタム データの操作の詳細については `Settings` [、「Persisting add-in state and settings」を参照してください](persisting-add-in-state-and-settings.md)。

## <a name="read-properties-of-a-project-document"></a>プロジェクト ドキュメントのプロパティの読み取り

作業ウィンドウ アドインが Project で動作する場合は、そのアドインでアクティブ プロジェクト内のプロジェクト フィールド、リソース、およびタスク フィールドの一部からデータを読み取ることができます。これを実現するには、追加の Project 固有の機能を提供するように [Document](/javascript/api/office/office.document) オブジェクトを拡張する `Document` オブジェクトのメソッドとイベントを使用します。

Project のデータの読み取り操作の例については、「[テキスト エディターを使用して Project 2013 用の作業ウィンドウ アドインを初めて作成する](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)」を参照してください。

## <a name="permissions-model-and-governance"></a>アクセス許可モデルとガバナンス

アドインは、マニフェスト内の要素を使用して、JavaScript API から必要な機能のレベルにアクセスするためのアクセス許可を要求Office `Permissions` します。 たとえば、アドインでドキュメントへの読み取り/書き込みアクセスが必要な場合、そのマニフェストは要素のテキスト値 `ReadWriteDocument` として指定する必要 `Permissions` があります。 アクセス許可はユーザーのプライバシーとセキュリティを保護するために存在しているので、ベスト プラクティスとしては、その機能に必要な最低限のアクセス許可を要求することをお勧めします。 次の例は、作業ウィンドウのマニフェストで **ReadDocument** アクセス許可を要求する方法を示しています。

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

詳細については、「アドイン [で API を使用するためのアクセス許可の要求」を参照してください](requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)。

## <a name="see-also"></a>関連項目

- [Office の JavaScript API](../reference/javascript-api-for-office.md)
- [Office アドイン マニフェストのスキーマ参照](../develop/add-in-manifests.md)
- [Office アドインでのユーザー エラーのトラブルシューティング](../testing/testing-and-troubleshooting.md)
