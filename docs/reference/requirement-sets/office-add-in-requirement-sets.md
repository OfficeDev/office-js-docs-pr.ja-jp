---
title: Office 共通 API の要件セット
description: ''
ms.date: 04/10/2019
ms.prod: non-product-specific
localization_priority: Priority
ms.openlocfilehash: 16f77e81d149aa2323760013f64fbf36f4ce7d8f
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450108"
---
# <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)」をご覧ください。

Office ホストによってアドインがサポートされる場所に関する情報が必要ですか? 「[Office アドインのホストとプラットフォームの可用性](/office/dev/add-ins/overview/office-add-in-availability)」を参照してください。

*ホスト固有*の API 要件セットをお探しですか? 次の API 要件セットを参照してください。

- [Excel JavaScript API 要件セット](excel-api-requirement-sets.md) (ExcelApi)
- [Word JavaScript API 要件セット](word-api-requirement-sets.md) (WordApi)
- [OneNote JavaScript API 要件セット](onenote-api-requirement-sets.md) (OneNoteApi)
- [Outlook API 要件セットについて](outlook-api-requirement-sets.md) (Mailbox)

> [!IMPORTANT]
> SharePoint で Access Web アプリとデータベースを作成して使用することは推奨されなくなりました。 代わりに、[Microsoft PowerApps](https://powerapps.microsoft.com/) を使用して、コード作成が不要な Web とモバイル デバイス用ビジネス ソリューションをビルドすることをお勧めします。

## <a name="common-api-requirement-sets"></a>共通 API の要件セット

次のセクションは、共通 API の要件セット、各セットのメソッド、その要件セットをサポートする Office ホスト アプリケーションの一覧です。これらの API 要件セットのバージョンはすべて 1.1 です。

### <a name="activeview"></a>ActiveView

|**Office のホスト**|**セット内のメソッド**|
|:-----|:-----|
| PowerPoint<br>PowerPoint Online<br>PowerPoint for iPad<br>PowerPoint for Mac|Document.getActiveViewAsync|

---

### <a name="addincommands"></a>AddInCommands

「[アドイン コマンドの要件セット](add-in-commands-requirement-sets.md)」を参照してください。

---

### <a name="bindingevents"></a>BindingEvents

|**Office のホスト**|**セット内のメソッド**|
|:-----|:-----|
| Access Web App<br>Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>Word 2013 以降<br>Word 2016 for Mac 以降<br>Word Online<br>Word for iPad|Binding.addHanderAsync<br>Binding.removeHanderAsync|

---

### <a name="compressedfile"></a>CompressedFile

|**Office のホスト**|**セット内のメソッド**|
|:-----|:-----|
| Excel<br>Excel Online<br>Excel for Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint for iPad<br>PowerPoint for Mac<br>Word 2013 以降<br>Word 2016 for Mac 以降<br>Word Online<br>Word for iPad|Document.getFileAsync メソッドを使用するときの、<br>バイト配列 (Office.FileType.Compressed) としての Office Open XML (OOXML) 形式への出力をサポートします。|

---

### <a name="customxmlparts"></a>CustomXmlParts

|**Office のホスト**|**セット内のメソッド**|
|:-----|:-----|
| Word 2013 以降<br>Word 2016 for Mac 以降<br>Word Online<br>Word for iPad|CustomXmlNode.getNodesAsync<br>CustomXmlNode.getNodeValueAsync<br>CustomXmlNode.getTextAsync<br>CustomXmlNode.getXmlAsync<br>CustomXmlNode.setNodeValueAsync<br>CustomXmlNode.setTextAsync<br>CustomXmlNode.setXmlAsync<br>CustomXmlPart.addHandlerAsync<br>CustomXmlPart.deleteAsync<br>CustomXmlPart.getNodesAsync<br>CustomXmlPart.getXmlAsync<br>CustomXmlPart.removeHandlerAsync<br>CustomXmlParts.addAsync<br>CustomXmlParts.getByIdAsync<br>CustomXmlParts.getByNamespaceAsync<br>CustomXmlPrefixMappings.addNamespaceAsync<br>CustomXmlPrefixMappings.getNamespaceAsync<br>CustomXmlPrefixMappings.getPrefixAsync|

---

### <a name="dialogapi"></a>DialogApi

|**Office のホスト**|**セット内のメソッド**|
|:-----|:-----|
| 「[ダイアログ API の要件セット](dialog-api-requirement-sets.md)」を参照してください。 | UI.messageParent<br>UI.displayDialogAsync<br>UI.closeContainer<br>UI.Dialog |

---

### <a name="documentevents"></a>DocumentEvents

|**Office のホスト**|**セット内のメソッド**|
|:-----|:-----|
| Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint for iPad<br>PowerPoint for Mac<br>Word 2013 以降<br>Word 2016 for Mac 以降<br>Word Online<br>Word for iPad|Document.addHandlerAsync<br>Document.removeHandlerAsync|

---

### <a name="file"></a>File

|**Office のホスト**|**セット内のメソッド**|
|:-----|:-----|
| Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint for iPad<br>PowerPoint for Mac<br>Word 2013 以降<br>Word 2016 for Mac 以降<br>Word Online<br>Word for iPad|Document.getFileAsync<br>File.closeAsync<br>File.getSliceAsync|

---

### <a name="htmlcoercion"></a>HtmlCoercion

|**Office のホスト**|**セット内のメソッド**|
|:-----|:-----|
| OneNote Online<br>Word 2013 以降<br>Word 2016 for Mac 以降<br>Word Online<br>Word for iPad|Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync、または Binding.setDataAsync メソッドを使用してデータを読み書きするときの、<br>HTML (Office.CoercionType.Html) への強制型変換をサポートします。|

---

### <a name="identityapi"></a>IdentityAPI

|**Office のホスト**|**セット内のメソッド**|
|:-----|:-----|
| 「[Identity API の要件セット](identity-api-requirement-sets.md)」を参照してください。 | Auth.getAccessTokenAsync |

---

### <a name="imagecoercion"></a>ImageCoercion

|**Office のホスト**|**セット内のメソッド**|
|:-----|:-----|
| Excel<br>Excel for iPad<br>Excel for Mac<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint for iPad<br>PowerPoint for Mac<br>Word 2013 以降<br>Word 2016 for Mac 以降<br>Word Online<br>Word for iPad|Document.setSelectedDataAsync メソッドを使用してデータを書き込むときに、画像 (Office.CoercionType.Image) への変換をサポートしています。|

---

### <a name="mailbox"></a>メールボックス

|**Office のホスト**|**セット内のメソッド**|
|:-----|:-----|
|Outlook for Windows<br>Outlook for web<br>Outlook for Android<br>Outlook for Mac<br>Outlook Web App |「[Outlook API 要件セットについて](outlook-api-requirement-sets.md)」をご覧ください。|

---

### <a name="matrixbindings"></a>MatrixBindings

|**Office のホスト**|**セット内のメソッド**|
|:-----|:-----|
| Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>Word<br>Word Online<br>Word for iPad<br>Word for Mac|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsync<br>Binding.getDataAsync<br>Binding.setDataAsync|

---

### <a name="matrixcoercion"></a>MatrixCoercion

|**Office のホスト**|**セット内のメソッド**|
|:-----|:-----|
| Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>Word 2013 以降<br>Word 2016 for Mac 以降<br>Word Online<br>Word for iPad|Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync、または Binding.setDataAsync の各メソッドを使用してデータを読み書きするとき、"matrix" (配列の配列) データ構造への強制型変換 (Office.CoercionType.Matrix) をサポートします。|

---

### <a name="ooxmlcoercion"></a>OoxmlCoercion

|**Office のホスト**|**セット内のメソッド**|
|:-----|:-----|
| Word 2013 以降<br>Word 2016 for Mac 以降<br>Word Online<br>Word for iPad|Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync、または Binding.setDataAsync の各メソッドを使用してデータを読み書きするとき、Open Office XML (OOXML) 形式への強制型変換 (Office.CoercionType.Ooxml) をサポートします。|

---

### <a name="partialtablebindings"></a>PartialTableBindings

|**Office のホスト**|**セット内のメソッド**|
|:-----|:-----|
| Access Web App||

---

### <a name="pdffile"></a>PdfFile

|**Office のホスト**|**セット内のメソッド**|
|:-----|:-----|
| Excel for Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint for iPad<br>PowerPoint for Mac<br>Word 2013 以降<br>Word 2016 for Mac 以降<br>Word Online<br>Word for iPad|Document.getFileAsync メソッドを使用するときの、<br>PDF 形式 (Office.FileType.Pdf) への出力をサポートします。|

---

### <a name="selection"></a>Selection

|**Office のホスト**|**セット内のメソッド**|
|:-----|:-----|
| Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint for iPad<br>PowerPoint for Mac<br>Project<br>Word 2013 以降<br>Word 2016 for Mac 以降<br>Word Online<br>Word for iPad|Document.getSelectedDataAsync<br>Document.setSelectedDataAsync|

---

### <a name="settings"></a>Settings

|**Office のホスト**|**セット内のメソッド**|
|:-----|:-----|
| Access Web App<br>Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint for iPad<br>PowerPoint for Mac<br>Word 2013 以降<br>Word 2016 for Mac 以降<br>Word Online<br>Word for iPad|Settings.get<br>Settings.remove<br>Settings.saveAsync<br>Settings.set|

---

### <a name="tablebindings"></a>TableBindings

|**Office のホスト**|**セット内のメソッド**|
|:-----|:-----|
| Access Web App<br>Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>Word 2013 以降<br>Word 2016 for Mac 以降<br>Word Online<br>Word for iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsync<br>Binding.addColumnsAsync<br>Binding.addRowsAsync<br>Binding.deleteAllDataValuesAsync<br>Binding.getDataAsync<br>Binding.setDataAsync|

---

### <a name="tablecoercion"></a>TableCoercion

|**Office のホスト**|**セット内のメソッド**|
|:-----|:-----|
| Access Web App<br>Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>Word 2013 以降<br>Word 2016 for Mac 以降<br>Word Online<br>Word for iPad|Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync、または Binding.setDataAsync の各メソッドを使用してデータを読み書きするとき、"table" データ構造への強制型変換 (Office.CoercionType.Table) をサポートします。|

---

### <a name="textbindings"></a>TextBindings

|**Office のホスト**|**セット内のメソッド**|
|:-----|:-----|
| Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>Word 2013 以降<br>Word 2016 for Mac 以降<br>Word Online<br>Word for iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsync<br>Binding.getDataAsync<br>Binding.setDataAsync|

---

### <a name="textcoercion"></a>TextCoercion

|**Office のホスト**|**セット内のメソッド**|
|:-----|:-----|
| Excel<br>Excel Online<br>Excel for iPad<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint for iPad<br>PowerPoint for Mac<br>Project<br>Word 2013 以降<br>Word 2016 for Mac 以降<br>Word Online<br>Word for iPad|Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync、または Binding.setDataAsync の各メソッドを使用してデータを読み書きするとき、テキスト形式への強制型変換 (Office.CoercionType.Text) をサポートします。|

---

### <a name="textfile"></a>TextFile

|**Office のホスト**|**セット内のメソッド**|
|:-----|:-----|
| Word 2013 以降<br>Word 2016 for Mac 以降<br>Word Online<br>Word for iPad|Document.getFileAsync メソッドを使用するとき、テキスト形式 (Office.FileType.Text) への出力をサポートします。|

---

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
