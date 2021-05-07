---
title: Office 共通 API の要件セット
description: 共通 API 要件セットOffice詳細を確認します。
ms.date: 04/28/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 959f03bf41496c1506087c2851efad336cdec676
ms.sourcegitcommit: 8fbc7c7eb47875bf022e402b13858695a8536ec5
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/06/2021
ms.locfileid: "52253348"
---
# <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office アプリケーションがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。

> [!TIP]
> アプリケーション固有の API *要件* セットをお探しの場合 次の API 要件セットを参照してください。
>
> - [Excel JavaScript API 要件セット](excel-api-requirement-sets.md) (ExcelApi)
> - [Word JavaScript API 要件セット](word-api-requirement-sets.md) (WordApi)
> - [OneNote JavaScript API 要件セット](onenote-api-requirement-sets.md) (OneNoteApi)
> - [PowerPoint JavaScript API 要件セット](powerpoint-api-requirement-sets.md) (PowerPointApi)
> - [Outlook API 要件セットについて](outlook-api-requirement-sets.md) (Mailbox)

> [!IMPORTANT]
> SharePoint で Access Web アプリとデータベースを作成して使用することは推奨されなくなりました。 代わりに、[Microsoft PowerApps](https://powerapps.microsoft.com/) を使用して、コード作成が不要な Web とモバイル デバイス用ビジネス ソリューションをビルドすることをお勧めします。

## <a name="common-api-requirement-sets"></a>共通 API の要件セット

以下のセクションでは、共通 API 要件セット、各セット内のメソッド、および要件セットをサポートするクライアント アプリケーションOffice一覧を示します。 特に指定がない限り、これらの API 要件セットのバージョンはすべて 1.1 です。

> [!TIP]
> アドインと要件セットがアプリケーションとバージョンでサポートされている場所Office必要ですか? [「Officeアドインのクライアント アプリケーションとプラットフォームの可用性Officeを参照してください](../../overview/office-add-in-availability.md)。

### <a name="activeview"></a>ActiveView

|**Office アプリケーション**|**セット内のメソッド**|
|:-----|:-----|
| PowerPoint on Windows<br>PowerPoint on the web<br>PowerPoint on iPad<br>PowerPoint on Mac|Document.getActiveViewAsync|

---

### <a name="addincommands"></a>AddInCommands

「[アドイン コマンドの要件セット](add-in-commands-requirement-sets.md)」を参照してください。

---

### <a name="bindingevents"></a>BindingEvents

|**Office アプリケーション**|**セット内のメソッド**|
|:-----|:-----|
| Access Web App<br>Excel on Windows<br>Excel on the web<br>Excel on iPad<br>Excel on Mac<br>Word on Windows (Word 2013 以降)<br>Word on Mac (Word 2016 以降)<br>Word on the web<br>Word on iPad|Binding.addHandlerAsync<br>Binding.removeHandlerAsync|

---

### <a name="compressedfile"></a>CompressedFile

|**Office アプリケーション**|**セット内のメソッド**|
|:-----|:-----|
| Excel 2016以降のWindows<br>Excel on the web<br>Excel 2016以降の Mac<br>PowerPoint on Windows<br>PowerPoint on the web<br>PowerPoint on iPad<br>PowerPoint on Mac<br>Word on Windows (Word 2013 以降)<br>Word on Mac (Word 2016 以降)<br>Word on the web<br>Word on iPad|Document.getFileAsync メソッドを使用するときの、<br>バイト配列 (Office.FileType.Compressed) としての Office Open XML (OOXML) 形式への出力をサポートします。|

---

### <a name="customxmlparts"></a>CustomXmlParts

|**Office アプリケーション**|**セット内のメソッド**|
|:-----|:-----|
| Word on Windows (Word 2013 以降)<br>Word on Mac (Word 2016 以降)<br>Word on the web<br>Word on iPad|CustomXmlNode.getNodesAsync<br>CustomXmlNode.getNodeValueAsync<br>CustomXmlNode.getTextAsync<br>CustomXmlNode.getXmlAsync<br>CustomXmlNode.setNodeValueAsync<br>CustomXmlNode.setTextAsync<br>CustomXmlNode.setXmlAsync<br>CustomXmlPart.addHandlerAsync<br>CustomXmlPart.deleteAsync<br>CustomXmlPart.getNodesAsync<br>CustomXmlPart.getXmlAsync<br>CustomXmlPart.removeHandlerAsync<br>CustomXmlParts.addAsync<br>CustomXmlParts.getByIdAsync<br>CustomXmlParts.getByNamespaceAsync<br>CustomXmlPrefixMappings.addNamespaceAsync<br>CustomXmlPrefixMappings.getNamespaceAsync<br>CustomXmlPrefixMappings.getPrefixAsync|

---

### <a name="dialogapi"></a>DialogApi

|**Office アプリケーション**|**セット内のメソッド**|
|:-----|:-----|
| 「[ダイアログ API の要件セット](dialog-api-requirement-sets.md)」を参照してください。 | UI.messageParent<br>UI.displayDialogAsync<br>UI.closeContainer<br>UI.Dialog |

---

### <a name="documentevents"></a>DocumentEvents

|**Office アプリケーション**|**セット内のメソッド**|
|:-----|:-----|
| Excel on Windows<br>Excel on the web<br>Excel on iPad<br>Excel on Mac<br>OneNote on the web<br>PowerPoint on Windows<br>PowerPoint on the web<br>PowerPoint on iPad<br>PowerPoint on Mac<br>Word on Windows (Word 2013 以降)<br>Word on Mac (Word 2016 以降)<br>Word on the web<br>Word on iPad|Document.addHandlerAsync<br>Document.removeHandlerAsync|

---

### <a name="file"></a>File

|**Office アプリケーション**|**セット内のメソッド**|
|:-----|:-----|
| Excel on Windows<br>Excel on the web<br>Excel on iPad<br>Excel on Mac<br>PowerPoint on Windows<br>PowerPoint on the web<br>PowerPoint on iPad<br>PowerPoint on Mac<br>Word on Windows (Word 2013 以降)<br>Word on Mac (Word 2016 以降)<br>Word on the web<br>Word on iPad|Document.getFileAsync<br>File.closeAsync<br>File.getSliceAsync|

---

### <a name="htmlcoercion"></a>HtmlCoercion

|**Office アプリケーション**|**セット内のメソッド**|
|:-----|:-----|
| OneNote on the web<br>Word on Windows (Word 2013 以降)<br>Word on Mac (Word 2016 以降)<br>Word on the web<br>Word on iPad|Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync、または Binding.setDataAsync の各メソッドを使用してデータを読み書きするとき、HTML への強制型変換 (Office.CoercionType.Html) をサポートします。|

---

### <a name="identityapi"></a>IdentityAPI

|**Office アプリケーション**|**セット内のメソッド**|
|:-----|:-----|
| 「[Identity API の要件セット](identity-api-requirement-sets.md)」を参照してください。 | Auth.getAccessToken |

---

### <a name="imagecoercion"></a>ImageCoercion

|**Office アプリケーション**|**セット内のメソッド**|
|:-----|:-----|
| 「[画像強制型変換要件セット](image-coercion-requirement-sets.md)」を参照してください。 | Document.setSelectedDataAsync メソッド|

---

### <a name="mailbox"></a>Mailbox

|**Office アプリケーション**|**セット内のメソッド**|
|:-----|:-----|
|Windows での Outlook<br>Outlook on the web<br>Outlook on Android<br>Outlook on Mac<br>Outlook on iOS|「[Outlook API 要件セットについて](outlook-api-requirement-sets.md)」をご覧ください。|

---

### <a name="matrixbindings"></a>MatrixBindings

|**Office アプリケーション**|**セット内のメソッド**|
|:-----|:-----|
| Excel on Windows<br>Excel on the web<br>Excel on iPad<br>Excel on Mac<br>Word on Windows<br>Word on the web<br>Word on iPad<br>Word on Mac|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsync<br>Binding.getDataAsync<br>Binding.setDataAsync|

---

### <a name="matrixcoercion"></a>MatrixCoercion

|**Office アプリケーション**|**セット内のメソッド**|
|:-----|:-----|
| Excel on Windows<br>Excel on the web<br>Excel on iPad<br>Excel on Mac<br>Word on Windows (Word 2013 以降)<br>Word on Mac (Word 2016 以降)<br>Word on the web<br>Word on iPad|Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync、または Binding.setDataAsync の各メソッドを使用してデータを読み書きするとき、"matrix" (配列の配列) データ構造への強制型変換 (Office.CoercionType.Matrix) をサポートします。|

---

### <a name="ooxmlcoercion"></a>OoxmlCoercion

|**Office アプリケーション**|**セット内のメソッド**|
|:-----|:-----|
| Word on Windows (Word 2013 以降)<br>Word on Mac (Word 2016 以降)<br>Word on the web<br>Word on iPad|Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync、または Binding.setDataAsync の各メソッドを使用してデータを読み書きするとき、Open Office XML (OOXML) 形式への強制型変換 (Office.CoercionType.Ooxml) をサポートします。|

---

### <a name="openbrowserwindowapi"></a>OpenBrowserWindowApi

|**Office のホスト**|**セット内のメソッド**|
|:-----|:-----|
| 「 [ブラウザー ウィンドウ API の要件セットを開く」を参照してください](open-browser-window-api-requirement-sets.md)。 | Office.context.ui.openBrowserWindow |

---

### <a name="partialtablebindings"></a>PartialTableBindings

|**Office アプリケーション**|**セット内のメソッド**|
|:-----|:-----|
| Access Web App||

---

### <a name="pdffile"></a>PdfFile

|**Office アプリケーション**|**セット内のメソッド**|
|:-----|:-----|
| Excel on Windows<br>Excel on the web<br>Excel on Mac<br>PowerPoint on Windows<br>PowerPoint on the web<br>PowerPoint on iPad<br>PowerPoint on Mac<br>Word on Windows (Word 2013 以降)<br>Word on Mac (Word 2016 以降)<br>Word on the web|Document.getFileAsync メソッドを使用するときの、<br>PDF 形式 (Office.FileType.Pdf) への出力をサポートします。|

---

### <a name="ribbonapi"></a>RibbonApi

|**Office アプリケーション**|**セット内のメソッド**|
|:-----|:-----|
| リボン [API 要件セットを参照してください](ribbon-api-requirement-sets.md)。 | Office.ribbon.requestUpdate |

---

### <a name="selection"></a>Selection

|**Office アプリケーション**|**セット内のメソッド**|
|:-----|:-----|
| Excel on Windows<br>Excel on the web<br>Excel on iPad<br>Excel on Mac<br>PowerPoint on Windows<br>PowerPoint on the web<br>PowerPoint on iPad<br>PowerPoint on Mac<br>Project on Windows<br>Word on Windows (Word 2013 以降)<br>Word on Mac (Word 2016 以降)<br>Word on the web<br>Word on iPad|Document.getSelectedDataAsync<br>Document.setSelectedDataAsync|

---

### <a name="settings"></a>Settings

|**Office アプリケーション**|**セット内のメソッド**|
|:-----|:-----|
| Access Web App<br>Excel on Windows<br>Excel on the web<br>Excel on iPad<br>Excel on Mac<br>OneNote on the web<br>PowerPoint on Windows<br>PowerPoint on the web<br>PowerPoint on iPad<br>PowerPoint on Mac<br>Word on Windows (Word 2013 以降)<br>Word on Mac (Word 2016 以降)<br>Word on the web<br>Word on iPad|Settings.get<br>Settings.remove<br>Settings.saveAsync<br>Settings.set|

---

### <a name="sharedruntime"></a>SharedRuntime

|**Office アプリケーション**|**セット内のメソッド**|
|:-----|:-----|
| 「共有 [ランタイム要件セット」を参照してください](shared-runtime-requirement-sets.md)。 | Office.addin.getStartupBehavior<br>Office.addin.hide<br>Office.addin.onVisibilityModeChanged<br>Office.addin.setStartupBehavior<br>Office.addin.showAsTaskpane<br> |

---

### <a name="tablebindings"></a>TableBindings

|**Office アプリケーション**|**セット内のメソッド**|
|:-----|:-----|
| Access Web App<br>Excel on Windows<br>Excel on the web<br>Excel on iPad<br>Excel on Mac<br>Word on Windows (Word 2013 以降)<br>Word on Mac (Word 2016 以降)<br>Word on the web<br>Word on iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsync<br>Binding.addColumnsAsync<br>Binding.addRowsAsync<br>Binding.deleteAllDataValuesAsync<br>Binding.getDataAsync<br>Binding.setDataAsync|

---

### <a name="tablecoercion"></a>TableCoercion

|**Office アプリケーション**|**セット内のメソッド**|
|:-----|:-----|
| Access Web App<br>Excel on Windows<br>Excel on the web<br>Excel on iPad<br>Excel on Mac<br>Word on Windows (Word 2013 以降)<br>Word on Mac (Word 2016 以降)<br>Word on the web<br>Word on iPad|Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync、または Binding.setDataAsync の各メソッドを使用してデータを読み書きするとき、"table" データ構造への強制型変換 (Office.CoercionType.Table) をサポートします。|

---

### <a name="textbindings"></a>TextBindings

|**Office アプリケーション**|**セット内のメソッド**|
|:-----|:-----|
| Excel on Windows<br>Excel on the web<br>Excel on iPad<br>Excel on Mac<br>Word 2013 以降および Windows<br>Word on Mac (Word 2016 以降)<br>Word on the web<br>Word on iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsync<br>Binding.getDataAsync<br>Binding.setDataAsync|

---

### <a name="textcoercion"></a>TextCoercion

|**Office アプリケーション**|**セット内のメソッド**|
|:-----|:-----|
| Excel on Windows<br>Excel on the web<br>Excel on iPad<br>OneNote on the web<br>PowerPoint on Windows<br>PowerPoint on the web<br>PowerPoint on iPad<br>PowerPoint on Mac<br>Project on Windows<br>Word on Windows (Word 2013 以降)<br>Word on Mac (Word 2016 以降)<br>Word on the web<br>Word on iPad|Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync、または Binding.setDataAsync の各メソッドを使用してデータを読み書きするとき、テキスト形式への強制型変換 (Office.CoercionType.Text) をサポートします。|

---

### <a name="textfile"></a>TextFile

|**Office アプリケーション**|**セット内のメソッド**|
|:-----|:-----|
| Word on Windows (Word 2013 以降)<br>Word on Mac (Word 2016 以降)<br>Word on the web<br>Word on iPad|Document.getFileAsync メソッドを使用するとき、テキスト形式 (Office.FileType.Text) への出力をサポートします。|

---

## <a name="methods-that-arent-part-of-a-requirement-set"></a>要件セットの一部ではないメソッド

JavaScript API の次Officeメソッドは、要件セットの一部ではない。 アドインでこれらのメソッドが必要な場合は、アドインのマニフェストで **Methods** 要素と **Method** 要素を使用してメソッドが必要であると宣言するか、または `if` ステートメントを使用してランタイム チェックを実行します。 詳細については、「アプリケーションと[API の要件Office指定する」を参照してください](../../develop/specify-office-hosts-and-api-requirements.md)。

|**メソッド名**|**Officeサポート**|
|:-----|:-----|
|Bindings.addFromPromptAsync|Access Web アプリ、Excel on Windows、Excel on the web、Excel on iPad、および Excel on Mac|
|Document.getFilePropertiesAsync|Excel on Windows、Excel on the web、Excel on iPad、Excel on Mac、PowerPoint on Windows、PowerPoint on the web、PowerPoint on iPad、PowerPoint on Mac、Word on Windows、Word on the web、Word on iPad、および Word on Mac|
|Document.getProjectFieldAsync|Project Standard 2013、Project Professional 2013|
|Document.getResourceFieldAsync|Project Standard 2013、Project Professional 2013|
|Document.getSelectedResourceAsync|Project Standard 2013、Project Professional 2013|
|Document.getSelectedTaskAsync|Project Standard 2013、Project Professional 2013|
|Document.getSelectedViewAsync|Project Standard 2013、Project Professional 2013|
|Document.getTaskAsync|Project Standard 2013、Project Professional 2013|
|Document.getTaskFieldAsync|Project Standard 2013、Project Professional 2013|
|Document.goToByIdAsync|Excel on Windows、Excel on the web、Excel on iPad、Excel on Mac、PowerPoint on Windows、PowerPoint on the web、PowerPoint on iPad、PowerPoint on Mac、Word on Windows、Word on the web、Word on iPad、および Word on Mac|
|Settings.addHandlerAsync|Access Web アプリおよび Excel on the web|
|Settings.refreshAsync|Access Web アプリ、Excel on Windows、Excel on the web、PowerPoint on Windows、PowerPoint on the web、Word、および Word on the web|
|Settings.removeHandlerAsync|Access Web アプリおよび Excel on the web|
|TableBinding.clearFormatsAsync|Excel on Windows、Excel on the web、Excel on iPad、および Excel on Mac|
|TableBinding.setFormatsAsync|Excel on Windows、Excel on the web、Excel on iPad、および Excel on Mac|
|TableBinding.setTableOptionsAsync|Excel on Windows、Excel on the web、Excel on iPad、および Excel on Mac|

## <a name="see-also"></a>関連項目

- [Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)
- [Office アプリケーションと API 要件を指定する](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office アドインの XML マニフェスト](../../develop/add-in-manifests.md)
