---
title: Office 共通 API の要件セット
description: ''
ms.date: 03/19/2019
ms.prod: non-product-specific
localization_priority: Priority
ms.openlocfilehash: 84eee3c085821e741f44fc4a413005cbc1a61951
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870199"
---
# <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)」をご覧ください。

Office ホストによってアドインがサポートされる場所に関する情報が必要ですか? 「[Office アドインのホストとプラットフォームの可用性](/office/dev/add-ins/overview/office-add-in-availability)」を参照してください。

*ホスト固有*の API 要件セットをお探しですか? 次の API 要件セットを参照してください。
 
- [Excel JavaScript API 要件セット](excel-api-requirement-sets.md) (ExcelApi)
- [Word JavaScript API 要件セット](word-api-requirement-sets.md) (WordApi)
- [OneNote JavaScript API 要件セット](onenote-api-requirement-sets.md) (OneNoteApi)
- [Outlook API 要件セットについて](outlook-api-requirement-sets.md) (MailBox)

> [!IMPORTANT]
> SharePoint で Access Web アプリとデータベースを作成して使用することは推奨されなくなりました。 代わりに、[Microsoft PowerApps](https://powerapps.microsoft.com/) を使用して、コード作成が不要な Web とモバイル デバイス用ビジネス ソリューションをビルドすることをお勧めします。

## <a name="common-api-requirement-sets"></a>共通 API の要件セット

次の表は、共通 API の要件セット、各セットのメソッド、その要件セットをサポートする Office ホスト アプリケーションの一覧です。これらの API 要件セットのバージョンはすべて 1.1 です。

|**要件セット**|**Office ホスト**|**セット内のメソッド**|
|:-----|:-----|:-----|
| ActiveView | PowerPoint<br>PowerPoint Online<br>PowerPoint for iPad<br>PowerPoint for Mac|Document.getActiveViewAsync|
| AddInCommands | 「[アドイン コマンドの要件セット](add-in-commands-requirement-sets.md)」を参照してください。 | |
| BindingEvents  | Access Web App<br>Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>Word 2013 以降<br>Word 2016 for Mac 以降<br>Word Online<br>Word for iPad|Binding.addHanderAsync<br>Binding.removeHanderAsync|
| CompressedFile    | Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint for iPad<br>PowerPoint for Mac<br>Word 2013 以降<br>Word 2016 for Mac 以降<br>Word Online<br>Word for iPad|Document.getFileAsync メソッドを使用するときの、<br>バイト配列 (Office.FileType.Compressed) としての Office Open XML (OOXML) 形式への出力をサポートします。|
| CustomXmlParts    | Word 2013 以降<br>Word 2016 for Mac 以降<br>Word Online<br>Word for iPad|CustomXmlNode.getNodesAsync<br>CustomXmlNode.getNodeValueAsync<br>CustomXmlNode.getTextAsync<br>CustomXmlNode.getXmlAsync<br>CustomXmlNode.setNodeValueAsync<br>CustomXmlNode.setTextAsync<br>CustomXmlNode.setXmlAsync<br>CustomXmlPart.addHandlerAsync<br>CustomXmlPart.deleteAsync<br>CustomXmlPart.getNodesAsync<br>CustomXmlPart.getXmlAsync<br>CustomXmlPart.removeHandlerAsync<br>CustomXmlParts.addAsync<br>CustomXmlParts.getByIdAsync<br>CustomXmlParts.getByNamespaceAsync<br>CustomXmlPrefixMappings.addNamespaceAsync<br>CustomXmlPrefixMappings.getNamespaceAsync<br>CustomXmlPrefixMappings.getPrefixAsync|
| DialogApi | 「[ダイアログ API の要件セット](dialog-api-requirement-sets.md)」を参照してください。 | UI.messageParent<br>UI.displayDialogAsync<br>UI.closeContainer<br>UI.Dialog |
| DocumentEvents    | Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint for iPad<br>PowerPoint for Mac<br>Word 2013 以降<br>Word 2016 for Mac 以降<br>Word Online<br>Word for iPad|Document.addHandlerAsync<br>Document.removeHandlerAsync|
| File  | Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint for iPad<br>PowerPoint for Mac<br>Word 2013 以降<br>Word 2016 for Mac 以降<br>Word Online<br>Word for iPad|Document.getFileAsync<br>File.closeAsync<br>File.getSliceAsync|
| HtmlCoercion  | OneNote Online<br>Word 2013 以降<br>Word 2016 for Mac 以降<br>Word Online<br>Word for iPad|Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync、または Binding.setDataAsync メソッドを使用してデータを読み書きするときの、<br>HTML (Office.CoercionType.Html) への強制型変換をサポートします。|
| IdentityAPI | 「[Identity API の要件セット](identity-api-requirement-sets.md)」を参照してください。 | Auth.getAccessTokenAsync |
| ImageCoercion | Excel<br>Excel for iPad<br>Excel for Mac<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint for iPad<br>PowerPoint for Mac<br>Word 2013 以降<br>Word 2016 for Mac 以降<br>Word Online<br>Word for iPad|Document.setSelectedDataAsync メソッドを使用してデータを書き込むときに、画像 (Office.CoercionType.Image) への変換をサポートしています。|
| メールボックス   |Outlook for Windows<br>Outlook for web<br>Outlook for Android<br>Outlook for Mac<br>Outlook Web App |「[Outlook API 要件セットについて](outlook-api-requirement-sets.md)」をご覧ください。|
| MatrixBindings    | Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>Word<br>Word Online<br>Word for iPad<br>Word for Mac|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncMatrix<br>Binding.getDataAsyncMatrix<br>Binding.setDataAsync|
| MatrixCoercion    | Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>Word 2013 以降<br>Word 2016 for Mac 以降<br>Word Online<br>Word for iPad|Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync、または Binding.setDataAsync の各メソッドを使用してデータを読み書きするとき、"matrix" (配列の配列) データ構造への強制型変換 (Office.CoercionType.Matrix) をサポートします。|
| OoxmlCoercion | Word 2013 以降<br>Word 2016 for Mac 以降<br>Word Online<br>Word for iPad|Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync、または Binding.setDataAsync の各メソッドを使用してデータを読み書きするとき、Open Office XML (OOXML) 形式への強制型変換 (Office.CoercionType.Ooxml) をサポートします。|
| PartialTableBindings  | Access Web App||
| PdfFile   | Excel for Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint for iPad<br>PowerPoint for Mac<br>Word 2013 以降<br>Word 2016 for Mac 以降<br>Word Online<br>Word for iPad|Document.getFileAsync メソッドを使用するときの、<br>PDF 形式 (Office.FileType.Pdf) への出力をサポートします。|
| Selection | Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint for iPad<br>PowerPoint for Mac<br>Project<br>Word 2013 以降<br>Word 2016 for Mac 以降<br>Word Online<br>Word for iPad|Document.getSelectedDataAsync<br>Document.setSelectedDataAsync|
| Settings  | Access Web App<br>Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint for iPad<br>PowerPoint for Mac<br>Word 2013 以降<br>Word 2016 for Mac 以降<br>Word Online<br>Word for iPad|Settings.get<br>Settings.remove<br>Settings.saveAsync<br>Settings.set|
| TableBindings | Access Web App<br>Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>Word 2013 以降<br>Word 2016 for Mac 以降<br>Word Online<br>Word for iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncTable<br>Binding.addColumnsAsyncTable<br>Binding.addRowsAsyncTable<br>Binding.deleteAllDataValuesAsyncTable<br>Binding.getDataAsyncTable<br>Binding.setDataAsync|
| TableCoercion | Access Web App<br>Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>Word 2013 以降<br>Word 2016 for Mac 以降<br>Word Online<br>Word for iPad|Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync、または Binding.setDataAsync の各メソッドを使用してデータを読み書きするとき、"table" データ構造への強制型変換 (Office.CoercionType.Table) をサポートします。|
| TextBindings  | Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>Word 2013 以降<br>Word 2016 for Mac 以降<br>Word Online<br>Word for iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncText<br>Binding.getDataAsyncText<br>Binding.setDataAsync|
| TextCoercion  | Excel<br>Excel Online<br>Excel for iPad<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint for iPad<br>PowerPoint for Mac<br>Project<br>Word 2013 以降<br>Word 2016 for Mac 以降<br>Word Online<br>Word for iPad|Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync、または Binding.setDataAsync の各メソッドを使用してデータを読み書きするとき、テキスト形式への強制型変換 (Office.CoercionType.Text) をサポートします。|
| TextFile  | Word 2013 以降<br>Word 2016 for Mac 以降<br>Word Online<br>Word for iPad|Document.getFileAsync メソッドを使用するとき、テキスト形式 (Office.FileType.Text) への出力をサポートします。|

## <a name="methods-that-arent-part-of-a-requirement-set"></a>要件セットの一部ではないメソッド

JavaScript API for Office の以下のメソッドは、要件セットの一部ではありません。 アドインでこれらのメソッドが必要な場合は、アドインのマニフェストで **Methods** 要素と **Method** 要素を使用してメソッドが必要であると宣言するか、または `if` ステートメントを使用してランタイム チェックを実行します。 詳細については、「[Office のホストと API の要件を指定する](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)」を参照してください。

|**メソッド名**|**サポートされる Office のホスト**|
|:-----|:-----|
|Bindings.addFromPromptAsync|Access Web アプリ、Excel、Excel Online、Excel for iPad、Excel for Mac|
|Document.getFilePropertiesAsync|Excel、Excel Online、Excel for iPad、Excel for Mac、PowerPoint、PowerPoint Online、PowerPoint for iPad、PowerPoint for Mac、Word、Word Online、Word for iPad、Word for Mac|
|Document.getProjectFieldAsync|Project Standard 2013、Project Professional 2013|
|Document.getResourceFieldAsync|Project Standard 2013、Project Professional 2013|
|Document.getSelectedResourceAsync|Project Standard 2013、Project Professional 2013|
|Document.getSelectedTaskAsync|Project Standard 2013、Project Professional 2013|
|Document.getSelectedViewAsync|Project Standard 2013、Project Professional 2013|
|Document.getTaskAsync|Project Standard 2013、Project Professional 2013|
|Document.getTaskFieldAsync|Project Standard 2013、Project Professional 2013|
|Document.goToByIdAsync|Excel、Excel Online、Excel for iPad、Excel for Mac、PowerPoint、PowerPoint Online、PowerPoint for iPad、PowerPoint for Mac、Word、Word Online、Word for iPad、Word for Mac|
|Settings.addHandlerAsync|Access Web アプリおよび Excel Online|
|Settings.refreshAsync|Access Web アプリ、Excel、Excel Online、PowerPoint、PowerPoint Online、Word、Word Online|
|Settings.removeHandlerAsync|Access Web アプリおよび Excel Online|
|TableBinding.clearFormatsAsync|Excel、Excel Online、Excel for iPad、Excel for Mac|
|TableBinding.setFormatsAsync|Excel、Excel Online、Excel for iPad、Excel for Mac|
|TableBinding.setTableOptionsAsync|Excel、Excel Online、Excel for iPad、Excel for Mac|

## <a name="see-also"></a>関連項目

- [Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Office のホストと API の要件を指定する](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office アドインの XML マニフェスト](/office/dev/add-ins/develop/add-in-manifests)
