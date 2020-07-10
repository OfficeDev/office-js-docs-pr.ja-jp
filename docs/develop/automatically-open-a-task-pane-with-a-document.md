---
title: ドキュメントで作業ウィンドウを自動的に開く
description: ドキュメントを開いたときに自動的に Office アドインを開くように構成する方法について説明します。
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: 85b421a569ccb83c3d07f0f10fd4767929332f96
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093708"
---
# <a name="automatically-open-a-task-pane-with-a-document"></a>ドキュメントで作業ウィンドウを自動的に開く

Office アドインでアドインコマンドを使用して、office アプリのリボンにボタンを追加することにより、Office UI を拡張することができます。 ユーザーがコマンド ボタンをクリックすると、アクション (作業ウィンドウを開くなど) が実行されます。

いくつかのシナリオでは、ドキュメントを開いたときに、ユーザーの明示的な操作なしで、自動的に作業ウィンドウを開くことが必要になります。 AddInCommands 1.1 要件セットに導入されている、作業ウィンドウの Autoopen 機能は、作業ウィンドウを自動的に開く必要があるシナリオで使用できます。


## <a name="how-is-the-autoopen-feature-different-from-inserting-a-task-pane"></a>Autoopen 機能と作業ウィンドウの挿入の相違点

ユーザーがアドイン コマンドを使用しないアドイン (Office 2013 で実行するアドインなど) を起動すると、それらはドキュメントに挿入され、そのドキュメントに永続化されます。 その結果として、別のユーザーがドキュメントを開くと、そのユーザーにアドインのインストールを求めるダイアログが表示され、作業ウィンドウが開きます。 このモデルの課題は、多くの場合、ユーザーがドキュメントでアドインを永続化したくないということです。 たとえば、Word ドキュメントで辞書アドインを使用する学生は、そのドキュメントを同級生や教師が開いたときに、アドインのインストールを求めるダイアログが表示されることを望まない場合もあります。

Autoopen 機能では、特定のドキュメントに特定の作業ウィンドウ アドインを永続化させるかどうかをユーザーが明示的に定義できます。

## <a name="support-and-availability"></a>サポートと可用性

autoopen 機能は現在 <!-- in **developer preview** and it is only --> 次の製品およびプラットフォームでサポートされています。

|**製品**|**プラットフォーム**|
|:-----------|:------------|
|<ul><li>Word</li><li>Excel</li><li>PowerPoint</li></ul>|すべての製品でサポートされているプラットフォーム: <ul><li>Office on Windows Desktop. Build 16.0.8121.1000+</li><li>Office on Mac. Build 15.34.17051500+</li><li>Office on the web</li></ul>|


## <a name="best-practices"></a>ベスト プラクティス

Autoopen 機能を使用するときには、次に示すベスト プラクティスを適用してください。

- Autoopen 機能は、アドイン ユーザーの作業効率の向上に役立つ場合に使用します。たとえば、次の場合に使用します。
  - When the document needs the add-in in order to function properly. For example, a spreadsheet that includes stock values that are periodically refreshed by an add-in. The add-in should open automatically when the spreadsheet is opened to keep the values up to date.
  - When the user will most likely always use the add-in with a particular document. For example, an add-in that helps users fill in or change data in a document by pulling information from a backend system.
- Allow users to turn on or turn off the autoopen feature. Include an option in your UI for users to choose to no longer automatically open the add-in task pane.  
- 要件セット検出を使用して autoopen 機能が使用可能かどうかを判断し、そうでない場合はフォールバック動作を提供します。
- アドインの使用率を人為的に増やすために、Autoopen 機能を使用しないでください。 特定のドキュメントでアドインを自動的に開くことが適切でない場合は、この機能によってユーザーに迷惑を持たせる可能性があります。

    > [!NOTE]
    > Microsoft では、Autoopen 機能の乱用を見つけた場合は、そのアドインを AppSource から排除することがあります。

- Don't use this feature to pin multiple task panes. You can only set one pane of your add-in to open automatically with a document.  

## <a name="implementation"></a>実装

Autoopen 機能を実装するには: 

- 自動的に開く作業ウィンドウを指定します。
- 作業ウィンドウを自動的に開くドキュメントにタグ設定します。

> [!IMPORTANT]
> The pane that you designate to open automatically will only open if the add-in is already installed on the user's device. If the user does not have the add-in installed when they open a document, the autoopen feature will not work and the setting will be ignored. If you also require the add-in to be distributed with the document you need to set the visibility property to 1; this can only be done using OpenXML, an example is provided later in this article.

### <a name="step-1-specify-the-task-pane-to-open"></a>手順 1: 開く作業ウィンドウを指定する

To specify the task pane to open automatically, set the [TaskpaneId](../reference/manifest/action.md#taskpaneid) value to **Office.AutoShowTaskpaneWithDocument**. You can only set this value on one task pane. If you set this value on multiple task panes, the first occurrence of the value will be recognized and the others will be ignored.

次の例では、Office.AutoShowTaskpaneWithDocument に設定された TaskPaneId の値を示しています。

```xml
<Action xsi:type="ShowTaskpane">
    <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
    <SourceLocation resid="Contoso.Taskpane.Url" />
</Action>
```

### <a name="step-2-tag-the-document-to-automatically-open-the-task-pane"></a>手順 2:作業ウィンドウを自動的に開くよう、ドキュメントにタグを設定する

You can tag the document to trigger the autoopen feature in one of two ways. Pick the alternative that works best for your scenario.  


#### <a name="tag-the-document-on-the-client-side"></a>クライアント側でドキュメントにタグを設定する

Office.js の [settings.set](/javascript/api/office/office.settings) メソッドを使用して、**Office.AutoShowTaskpaneWithDocument** を **true** に設定します。次に例を示します。

```js
Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
Office.context.document.settings.saveAsync();
```

このメソッドは、アドインの対話式操作の一環としてドキュメントにタグを設定する必要がある場合に使用します (たとえば、ユーザーがバインディングを作成した直後に、または自動的にウィンドウを開くことを示すオプションを選択した直後に使用します)。

#### <a name="use-open-xml-to-tag-the-document"></a>Open XML を使用してドキュメントにタグを設定する

You can use Open XML to create or modify a document and add the appropriate Open Office XML markup to trigger the autoopen feature. For a sample that shows you how to do this, see [Office-OOXML-EmbedAddin](https://github.com/OfficeDev/Office-OOXML-EmbedAddin).

次に示す 2 つの Open XML パートをドキュメントに追加します。

- `webextension` パート
- `taskpane` パート

次の例は、`webextension` パートを追加する方法を示しています。

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

`webextension` パートには、プロパティ バッグと **Office.AutoShowTaskpaneWithDocument** という名前のプロパティが含まれています。このプロパティは、`true` に設定する必要があります。

また、`webextension` パートには、属性が `id`、`storeType`、`store`、および `version` のストアまたはカタログへの参照も含まれています。 Autoopen 機能に関連する `storeType` の値は、4 つのみです。 その他の 3 つの属性の値は、次の表に示すように、`storeType` の値に応じて決まります。

| **`storeType` 値** | **`id` 値**    |**`store` 値** | **`version` 値**|
|:---------------|:---------------|:---------------|:---------------|
|OMEX (AppSource)|アドインの AppSource アセット ID (注を参照)|AppSource のロケール (たとえば、"en-us")。|AppSource カタログのバージョン (注を参照)|
|FileSystem (ネットワーク共有)|アドイン マニフェストでのアドインの GUID。|ネットワーク共有のパス。例: "\\\\MyComputer\\MySharedFolder"。|アドイン マニフェストでのバージョン。|
|EXCatalog (Exchange サーバー経由の展開) |アドイン マニフェストでのアドインの GUID。|"EXCatalog"。 EXCatalog 行は、Microsoft 365 管理センターで一元展開を使用するアドインで使用する行です。|アドイン マニフェストでのバージョン。
|Registry (システム レジストリ)|アドイン マニフェストでのアドインの GUID。|"developer"|アドイン マニフェストでのバージョン。|

> [!NOTE]
> To find the asset ID and version of an add-in in AppSource, go to the AppSource landing page for the add-in. The asset ID appears in the address bar in the browser. The version is listed in the **Details** section of the page.

webextension マークアップの詳細については、「[[MS-OWEXML] 2.2.5.WebExtensionReference](https://msdn.microsoft.com/library/hh695383(v=office.12).aspx)」を参照してください。

次の例は、`taskpane` パートを追加する方法を示しています。

```xml
<wetp:taskpane dockstate="right" visibility="0" width="350" row="4" xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">
  <wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" />
</wetp:taskpane>
```

この例では、`visibility` 属性が "0" に設定されている点に注目してください。 これは、webextension パートと `taskpane` パートの追加後に、初めてドキュメントを開いたときに、ユーザーはリボンの **[アドイン]** ボタンからアドインをインストールする必要があることを意味します。 それ以降は、ファイルを開いたときに、アドイン作業ウィンドが自動的に開きます。 また、`visibility` を "0" に設定すると、ユーザーが Autoopen 機能をオン/オフできるようにするために Office.js を使用できるようにもなります。 具体的には、スクリプトでドキュメント設定の **Office.AutoShowTaskpaneWithDocument** を `true` または `false` に設定します  (詳細については、「[クライアント側でドキュメントにタグを設定する](#tag-the-document-on-the-client-side)」を参照してください)。

If `visibility` is set to "1", the task pane opens automatically the first time the document is opened. The user is prompted to trust the add-in, and when trust is granted, the add-in opens. Thereafter, the add-in task pane opens automatically when the file is opened. However, when `visibility` is set to "1", you can't use Office.js to enable users to turn on or turn off the autoopen feature.

アドインとドキュメントのテンプレートまたはコンテンツが緊密に統合されていて、ユーザーが Autoopen 機能をオフにすることない場合は、`visibility` を "1" に設定することが適切な選択になります。

> [!NOTE]
> If you want to distribute your add-in with the document, so that users are prompted to install it, you must set the visibility property to 1. You can only do this via Open XML.

An easy way to write the XML is to first run your add-in and [tag the document on the client side](#tag-the-document-on-the-client-side) to write the value, and then save the document and inspect the XML that is generated. Office will detect and provide the appropriate attribute values. You can also use the [Open XML SDK 2.5 Productivity Tool](https://www.microsoft.com/download/details.aspx?id=30425) tool to generate C# code to programmatically add the markup based on the XML you generate.

## <a name="test-and-verify-opening-task-panes"></a>作業ウィンドウ表示のテストと検証

Microsoft 365 管理センターを介して一元的な展開を使用して作業ウィンドウを自動的に開くように、アドインのテストバージョンを展開することができます。 次の例では、EXCatalog のストア版を使用して一元展開カタログからアドインを挿入する方法を示します。

```xml
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{52811C31-4593-43B8-A697-EB873422D156}">
    <we:reference id="af8fa5ba-4010-4bcc-9e03-a91ddadf6dd3" version="1.0.0.0" store="EXCatalog" storeType="EXCatalog"/>
    <we:alternateReferences/>
    <we:properties/>
    <we:bindings/>
    <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>
```

前の例をテストするには、Microsoft 365 サブスクリプションを使用して一元展開を試行し、アドインが想定どおりに動作することを確認します。 Microsoft 365 サブスクリプションをまだお持ちでない場合は、 [microsoft 365 開発者プログラム](https://developer.microsoft.com/office/dev-program)に参加することによって、更新可能な90日間の microsoft 365 サブスクリプションを無料で入手できます。

## <a name="see-also"></a>関連項目

Autoopen 機能の使用方法を示すサンプルについては、「[Office-Add-in-Commands-Samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/AutoOpenTaskpane)」を参照してください。
[Microsoft 365 開発者プログラムに参加](/office/developer-program/office-365-developer-program)します。
