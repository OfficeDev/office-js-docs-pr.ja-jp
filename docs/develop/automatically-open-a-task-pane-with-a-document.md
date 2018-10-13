---
title: ドキュメントで作業ウィンドウを自動的に開く
description: ''
ms.date: 05/02/2018
ms.openlocfilehash: 2ebce1ce8bd95ee7802b5509d375f1986bb2877e
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505917"
---
# <a name="automatically-open-a-task-pane-with-a-document"></a>ドキュメントで作業ウィンドウを自動的に開く

Office アドインでアドイン コマンドを使用すると、Office リボンにボタンを追加して Office UI を拡張できます。ユーザーがコマンド ボタンをクリックすると、アクション (作業ウィンドウを開くなど) が実行されます。 

一部のシナリオでは、ドキュメントを開いたときに、ユーザーが明示的に操作を行うことなく、自動的に作業ウィンドウを開く必要があります。AddInCommands 1.1 要件セットに導入されている作業ウィンドウの Autoopen 機能は、作業ウィンドウを自動的に開く必要があるシナリオで使用できます。 


## <a name="how-is-the-autoopen-feature-different-from-inserting-a-task-pane"></a>Autoopen 機能と作業ウィンドウの挿入の相違点 

ユーザーがアドイン コマンドを使用しないアドイン (Office 2013 で実行するアドインなど) を起動すると、そのアドインはドキュメントに挿入され保持されます。その結果、別のユーザーがドキュメントを開くと、そのユーザーにアドインのインストールを求めるダイアログが表示され作業ウィンドウが開きます。このモデルの問題点は、多くの場合においてユーザーの意に反してドキュメントにアドインが保持されてしまうことです。たとえば、Word ドキュメントで辞書アドインを使用する学生は、そのドキュメントを同級生や教師が開いたときにアドインのインストールを求めるダイアログが表示されることを望まない場合もあります。  

Autoopen 機能では、特定のドキュメントに特定の作業ウィンドウ アドインを保持させるかどうかをユーザーが明示的に定義できます。 

## <a name="support-and-availability"></a>サポートと可用性
現時点では、Autoopen 機能は次の製品およびプラットフォームで<!-- in **developer preview** and it is only -->サポートされています。

|**製品**|**プラットフォーム**|
|:-----------|:------------|
|<ul><li>Word</li><li>Excel</li><li>PowerPoint</li></ul>|すべての製品でサポートされているプラ​​ットフォーム。<ul><li>Windows デスクトップ版 Office (ビルド 16.0.8121.1000 以降)</li><li>Office for Mac (ビルド 15.34.17051500 以降)</li><li>Office Online</li></ul>|


## <a name="best-practices"></a>ベスト プラクティス

Autoopen 機能を使用する際には、次に示すベスト プラクティスを適用してください。

- Autoopen 機能は、アドイン ユーザーの作業効率の向上に役立つ場合に使用します。いくつかの例を示します。
    - ドキュメントが適切に機能するためにアドインを必要とする場合。たとえば、アドインで定期的に最新の株価を更新するスプレッドシートでは、最新の値を維持するためにスプレッドシートが開かれたときにアドインが自動的に開かれる必要があります。 
    - ユーザーが特定のドキュメントで常にアドインを使用する可能性が高い場合。たとえば、バックエンド システムから情報を取得してドキュメントのデータを入力または変更することでユーザーを支援するアドインです。 
- Autoopen 機能をユーザーがオン/オフできるようにします。ユーザーの UI に、アドインの作業ウィンドウが自動的に起動されないようにするオプションを含めます。  
- 要件セットの検出を使用して Autoopen 機能が利用可能かどうかを確認し、利用できない場合のフォールバック処理を用意します。
- アドインの使用率を人為的に増やすために、Autoopen 機能を使用しないでください。特定のドキュメントでアドインが意味もなく自動的に起動することはユーザーの妨げになります。 

    > [!NOTE]
    > Microsoft が Autoopen 機能の乱用を見つけた場合には、そのアドインを AppSource から排除する可能性があります。 

- この機能は、複数の作業ウィンドウを固定するために使用しないでください。1 つのドキュメントで自動的に開くアドインのウィンドウは 1 つのみ設定できます。  

## <a name="implementation"></a>実装
Autoopen 機能は次のように実装します。

- 自動的に開く作業ウィンドウを指定します。
- 作業ウィンドウを自動的に開くドキュメントにタグを設定します。

> [!IMPORTANT]
> 自動的に開くように指定したウィンドウは、アドインがユーザーのデバイスに既にインストールされている場合にのみ開きます。ユーザーがドキュメントを開いたときに、アドインがインストールされていない場合、Autoopen 機能は動作せず、設定は無視されます。また、アドインをドキュメントと共に配布する必要がある場合は、可視性プロパティを 1 に設定する必要があります。これは、OpenXML を使用する場合にのみ実行できます。例については、この記事で後述します。 

### <a name="step-1-specify-the-task-pane-to-open"></a>手順 1: 開く作業ウィンドウを指定する
自動的に開く作業ウィンドウを指定するには、[TaskpaneId](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/action?view=office-js#taskpaneid) の値を **Office.AutoShowTaskpaneWithDocument** に設定します。この値は 1 つの作業ウィンドウにのみ設定できます。この値を複数の作業ウィンドウに設定すると、最初に見つかった値のみを認識し、その他は無視されます。 

次の例は、TaskPaneId の値を Office.AutoShowTaskpaneWithDocument に設定しています。
          
```xml
<Action xsi:type="ShowTaskpane">
    <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
    <SourceLocation resid="Contoso.Taskpane.Url" />
</Action>
```     

### <a name="step-2-tag-the-document-to-automatically-open-the-task-pane"></a>手順 2: 作業ウィンドウを自動的に開くドキュメントにタグを設定する

ドキュメントにタグを設定し Autoopen 機能をトリガーする方法は 2 つあります。シナリオに最も適した方法を選択してください。  


#### <a name="tag-the-document-on-the-client-side"></a>クライアント側でドキュメントにタグを設定する
Office.js の [settings.set](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js) メソッドを使用して、**Office.AutoShowTaskpaneWithDocument** を **true** に設定します。次に例を示します。   

```js
Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
Office.context.document.settings.saveAsync();
```

このメソッドは、アドインの対話型操作の一部としてドキュメントにタグを設定する必要がある場合に使用します (たとえば、ユーザーがバインディングを作成した直後、または自動的にウィンドウを開くことを示すオプションを選択した直後にウィンドウを開く場合に使用します)。

#### <a name="use-open-xml-to-tag-the-document"></a>Open XML を使用してドキュメントにタグを設定する
Open XML を使用して、ドキュメントを作成または変更し、Autoopen 機能をトリガーするために必要な Open Office XML マークアップを追加できます。この方法を示すサンプルについては、「[Office-OOXML-EmbedAddin](https://github.com/OfficeDev/Office-OOXML-EmbedAddin)」を参照してください。 

次の 2 つの Open XML パートをドキュメントに追加します。

- webextension パート
- taskpane パート

次の例は、webextension パートを追加する方法を示します。

```xml
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="[ADD-IN ID PER MANIFEST]">
  <we:reference id="[GUID or AppSource asset ID]" version="[your add-in version]" store="[Pointer to store or catalog]" storeType="[Store or catalog type]"/>
  <we:alternateReferences/>
  <we:properties>
    <we:property name="Office.AutoShowTaskpaneWithDocument" value="true"/>
  </we:properties>
  <we:bindings/>
  <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>
```

webextension パートには、プロパティ バッグと **Office.AutoShowTaskpaneWithDocument** という名前のプロパティが含まれています。このプロパティは、`true` に設定する必要があります。

また、webextension パートには、ストアまたはカタログへの参照となる `id`、`storeType`、`store`、`version` の属性も含まれています。`storeType` 属性が Autoopen 機能に関連して持つ値は 4 つのみです。その他の 3 つの属性の値は、次の表に示すように、`storeType` の値に応じて決まります。 

| **`storeType` 値** | **`id` 値**    |**`store` 値** | **`version` 値**|
|:---------------|:---------------|:---------------|:---------------|
|OMEX (AppSource)|アドインの AppSource 資産 ID (注を参照)|AppSource のロケール (例: "en-us")|AppSource カタログのバージョン (注を参照)|
|FileSystem (ネットワーク共有)|アドイン マニフェスト内のアドインの GUID|ネットワーク共有のパス。例: "\\\\MyComputer\\MySharedFolder"|アドイン マニフェスト内のバージョン|
|EXCatalog (Exchange Server 経由の展開) |アドイン マニフェスト内のアドインの GUID|"EXCatalog"。EXCatalog 行は、Office 365 管理センターで一元展開を使用するアドインで使用する行です|アドイン マニフェスト内のバージョン
|Registry (システム レジストリ)|アドイン マニフェスト内のアドインの GUID|"developer"|アドイン マニフェスト内のバージョン|

> [!NOTE]
> AppSource でのアドインの資産 ID とバージョンを確認するには、そのアドインの AppSource ランディング ページに移動します。資産 ID は、ブラウザのアドレス バーに表示されます。バージョンは、そのページの **[詳細]** セクションに示されます。

webextension マークアップの詳細については、「[[MS-OWEXML] 2.2.5.WebExtensionReference](https://msdn.microsoft.com/library/hh695383(v=office.12).aspx)」を参照してください。

次の例は、taskpane パートを追加する方法を示しています。

```xml
<wetp:taskpane dockstate="right" visibility="0" width="350" row="4" xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">
  <wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" />
</wetp:taskpane>
```

この例では、`visibility` 属性が "0" に設定されている点に注目してください。これは、webextension パートと taskpane パートの追加後に、初めてドキュメントを開いたときに、ユーザーはリボンの **[アドイン]** ボタンからアドインをインストールする必要があることを意味します。それ以降は、ファイルを開くとアドイン作業ウィンドが自動的に開きます。また、`visibility` を "0" に設定すると、ユーザーが Autoopen 機能をオン/オフできるようにするために Office.js を使用できるようにもなります。具体的には、スクリプトでドキュメント設定の **Office.AutoShowTaskpaneWithDocument** を `true` または `false` に設定します (詳細については、「[クライアント側でドキュメントにタグを設定する](#tag-the-document-on-the-client-side)」を参照してください)。 

`visibility` が "1" に設定されていると、初めてドキュメントを開いたときに、自動的に作業ウィンドウが開きます。アドインを信頼することを求めるダイアログがユーザーに表示され、信頼が付与されるとアドインが開きます。それ以降は、ファイルを開くとアドイン作業ウィンドが自動的に開きます。ただし、`visibility` を "1" に設定すると、ユーザーが Autoopen 機能をオン/オフできるようにするために Office.js を使用することができなくなります。 

アドインとドキュメントのテンプレートまたはコンテンツが緊密に統合されているために、ユーザーが Autoopen 機能をオフにすることがない場合には、`visibility` を "1" に設定することが適切な選択になります。 

> [!NOTE]
> ドキュメントとともに配布するアドインのインストールをユーザーに求めるためには、visibility プロパティを 1 に設定する必要があります。これは、Open XML でのみ実行できます。

この XML を記述する簡単な方法は、最初にアドインを実行し、値を書き込むために[クライアント側でドキュメントにタグを設定](#tag-the-document-on-the-client-side)し、ドキュメントを保存してから生成された XML を調べる方法です。この方法では、Office は適切な属性値を検出し設定します。また、[Open XML SDK 2.5 Productivity Tool](https://www.microsoft.com/download/details.aspx?id=30425) を使用して生成した C# コードにより、生成する XML  に基づくプログラムでマークアップを追加することもできます。

## <a name="test-and-verify-opening-taskpanes"></a>作業ウィンドウ表示のテストと検証
自動的に作業ウィンドウを開くアドインのテスト バージョンは、Office 365 管理センターによる一元展開を使用して展開できます。次の例は、EXCatalog のストア版を使用して一元展開カタログからアドインを挿入する方法を示すものです。

```xml
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{52811C31-4593-43B8-A697-EB873422D156}">
    <we:reference id="af8fa5ba-4010-4bcc-9e03-a91ddadf6dd3" version="1.0.0.0" store="EXCatalog" storeType="EXCatalog"/>
    <we:alternateReferences/>
    <we:properties/>
    <we:bindings/>
    <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>
```
前の例をテストするために、[Office 365 開発者プログラム](https://docs.microsoft.com/office/developer-program/office-365-developer-program) に参加し、Office 365 サブスクリプションを購入していない場合は、[Office 365 開発者アカウント](https://developer.microsoft.com/office/dev-program) にサインアップすることを検討してください。実際に一元展開をテストし、アドインが期待どおりに動作することを確認できます。


## <a name="see-also"></a>関連項目

Autoopen 機能の使用方法を示すサンプルについては、 [Office アドイン コマンドのサンプル](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/AutoOpenTaskpane)を参照してください。 [Office 365 開発者プログラムに参加](https://docs.microsoft.com/office/developer-program/office-365-developer-program)します。 

