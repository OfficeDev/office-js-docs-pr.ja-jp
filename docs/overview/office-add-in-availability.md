---
title: Office アドインの Office クライアント アプリケーションとプラットフォームの可用性
description: Excel、OneNote、Outlook、PowerPoint、Project、Word のサポートされる要件セット。
ms.date: 05/11/2021
localization_priority: Priority
ms.openlocfilehash: ea41a8116bb1e54f9e7dcb290e1a48c61e6f0f68
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590870"
---
# <a name="office-client-application-and-platform-availability-for-office-add-ins"></a>Office アドインの Office クライアント アプリケーションとプラットフォームの可用性

期待どおりの動作をするうえで、Office アドインは特定の Office アプリケーション、要件セット、API メンバー、または API のバージョンに依存することがあります。次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、および各 Office アプリケーションで現在サポートされている共通 API が含まれています。

<br>

|<a href="#excel"><img src="../images/index/logo-excel.svg" alt="Excel" width="48" /><br><span>Excel</span></a>|<a href="#onenote"><img src="../images/index/logo-onenote.svg" alt="OneNote" width="48" /><br><span>OneNote</span></a>|<a href="#outlook"><img src="../images/index/logo-outlook.svg" alt="Outlook" width="48" /><br><span>Outlook</span></a>|<a href="#powerpoint"><img src="../images/index/logo-powerpoint.svg" alt="PowerPoint" width="48" /><br><span>PowerPoint</span></a>|<a href="#project"><img src="../images/index/logo-project-server.svg" alt="Project" width="48" /><br><span>Project</span></a>|<a href="#word"><img src="../images/index/logo-word.svg" alt="Word" width="48" /><br><span>Word</span></a>|
|:---:|:---:|:---:|:---:|:---:|:---:|

> [!NOTE]
> MSI からインストールされる最初の Office 2016 リリースには、ExcelApi 1.1、WordApi 1.1、共通 API の要件セットのみが含まれています。 さまざまなバージョンの Office の更新履歴の詳細については、「[関連項目](#see-also)」セクションをご確認ください。

## <a name="excel"></a>Excel

<table style="width:80%">
  <tr>
    <th style="width:10%">プラットフォーム</th>
    <th style="width:10%">拡張点</th>
    <th style="width:20%">API 要件セット</th>
    <th style="width:40%"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></th>
  </tr>
  <tr>
    <td>Office on the web</td>
    <td>
      - 作業ウィンドウ<br>
      - コンテンツ<br>
      - CustomFunctions<br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-12-requirement-set">ExcelApi 1.12</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets">IdentityAPI 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/ribbon-api-requirement-sets">RibbonAPI 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/shared-runtime-requirement-sets">SharedRuntime 1.1</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">設定</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Windows での Office<br>(Microsoft 365 サブスクリプションに接続)</td>
    <td>
      - 作業ウィンドウ<br>
      - コンテンツ<br>
      - CustomFunctions<br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-12-requirement-set">ExcelApi 1.12</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets">IdentityAPI 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/open-browser-window-api-requirement-sets">OpenBrowserWindowApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/ribbon-api-requirement-sets">RibbonAPI 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/shared-runtime-requirement-sets">SharedRuntime 1.1</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">設定</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Windows 版 Office 2019<br>(1 回限りの購入)</td>
    <td>
      - 作業ウィンドウ<br>
      - コンテンツ<br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">設定</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Windows 版 Office 2016<br>(1 回限りの購入)</td>
    <td>
      - 作業ウィンドウ<br>
      - コンテンツ </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*<br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">設定</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Windows 版 Office 2013<br>(1 回限りの購入)</td>
    <td>
      - 作業ウィンドウ<br>
      - コンテンツ </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*<br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">設定</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office on iPad<br>(Microsoft 365 サブスクリプションに接続)</td>
    <td>
      - 作業ウィンドウ<br>
      - コンテンツ </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-12-requirement-set">ExcelApi 1.12</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/open-browser-window-api-requirement-sets">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">設定</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office on Mac<br>(Microsoft 365 サブスクリプションに接続)</td>
    <td>
      - 作業ウィンドウ<br>
      - コンテンツ<br>
      - CustomFunctions<br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-12-requirement-set">ExcelApi 1.12</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets">IdentityAPI 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/open-browser-window-api-requirement-sets">OpenBrowserWindowApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/ribbon-api-requirement-sets">RibbonAPI 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/shared-runtime-requirement-sets">SharedRuntime 1.1</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">設定</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Mac 上の Office 2019<br>(1 回限りの購入)</td>
    <td>
      - 作業ウィンドウ<br>
      - コンテンツ<br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">設定</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Mac 上の Office 2016<br>(1 回限りの購入)</td>
    <td>
      - 作業ウィンドウ<br>
      - コンテンツ </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*<br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">設定</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </td>
  </tr>
</table>

*&ast; - リリース後の更新プログラムで追加されました。*

## <a name="custom-functions-excel-only"></a>カスタム関数 (Excel のみ)

<table style="width:80%">
  <tr>
    <th>プラットフォーム</th>
    <th>拡張点</th>
    <th>API 要件セット</th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></th>
  </tr>
  <tr>
    <td>Office on the web</td>
    <td>- CustomFunctions</td>
    <td>
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a><br>
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.2</a><br>
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.3</a>
    </td>
    <td></td>
  </tr>
  <tr>
    <td>Windows での Office<br>(Microsoft 365 サブスクリプションに接続)</td>
    <td>- CustomFunctions</td>
    <td>
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a><br>
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.2</a><br>
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.3</a>
    </td>
    <td></td>
  </tr>
  <tr>
    <td>Office on Mac<br>(Microsoft 365 サブスクリプションに接続)</td>
    <td>- CustomFunctions</td>
    <td>
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a><br>
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.2</a><br>
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.3</a>
    </td>
    <td></td>
  </tr>
</table>

## <a name="outlook"></a>Outlook

<table style="width:80%">
  <tr>
    <th>プラットフォーム</th>
    <th>拡張点</th>
    <th>API 要件セット</th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></th>
  </tr>
  <tr>
    <td>Office on the web<br>(モダン)</td>
    <td>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9">Mailbox 1.9</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10">Mailbox 1.10</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets">IdentityAPI 1.3</a><sup>1</sup>
    </td>
    <td>利用不可</td>
  </tr>
  <tr>
    <td>Office on the web<br>(クラシック)</td>
    <td>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a>
    </td>
    <td>利用不可</td>
  </tr>
  <tr>
    <td>Windows での Office<br>(Microsoft 365 サブスクリプションに接続)</td>
    <td>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">モジュール</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9">Mailbox 1.9</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10">Mailbox 1.10</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets">IdentityAPI 1.3</a><sup>1</sup><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/open-browser-window-api-requirement-sets">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>利用不可</td>
  </tr>
  <tr>
    <td>Windows 版 Office 2019<br>(1 回限りの購入)</td>
    <td>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">モジュール</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a>
    </td>
    <td>利用不可</td>
  </tr>
  <tr>
    <td>Windows 版 Office 2016<br>(1 回限りの購入)</td>
    <td>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">モジュール</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><sup>2</sup>
    </td>
    <td>使用不可</td>
  </tr>
  <tr>
    <td>Windows 版 Office 2013<br>(1 回限りの購入)</td>
    <td>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a><br>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><sup>2</sup><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><sup>2</sup>
    </td>
    <td>使用不可</td>
  </tr>
  <tr>
    <td>iOS 上の Office<br>(Microsoft 365 サブスクリプションに接続)</td>
    <td>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">既読メッセージ</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface">予定の開催者 (作成): オンライン会議</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a>
    </td>
    <td>利用不可</td>
  </tr>
  <tr>
    <td>Office on Mac<br>(現在の UI、<br>Microsoft 365 サブスクリプションに接続)</td>
    <td>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets">IdentityAPI 1.3</a><sup>1</sup><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/open-browser-window-api-requirement-sets">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>利用不可</td>
  </tr>
  <tr>
    <td>Office on Mac<br>(新しい UI (プレビュー)<sup>3</sup><br>Microsoft 365 サブスクリプションに接続)</td>
    <td>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets">IdentityAPI 1.3</a><sup>1</sup>
    </td>
    <td>利用不可</td>
  </tr>
  <tr>
    <td>Mac 上の Office 2019<br>(1 回限りの購入)</td>
    <td>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a>
    </td>
    <td>利用不可</td>
  </tr>
  <tr>
    <td>Mac 上の Office 2016<br>(1 回限りの購入)</td>
    <td>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a>
    </td>
    <td>利用不可</td>
  </tr>
  <tr>
    <td>Android 上の Office<br>(Microsoft 365 サブスクリプションに接続)</td>
    <td>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">既読メッセージ</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface">予定の開催者 (作成): オンライン会議</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a>
    </td>
    <td>利用不可</td>
  </tr>
</table>

> [!NOTE]
> <sup>1</sup> アドイン コードで Identity API セット 1.3 を要求するには、`isSetSupported('IdentityAPI', '1.3')` を呼び出してサポートされているかどうかを確認します。 アドイン マニフェストでの宣言はサポートされていません。 `undefined` ではないことを確認することで、API がサポートされているかどうかを判断することもできます。 詳細については、「[後続の要件セットからの API の使用](../reference/requirement-sets/outlook-api-requirement-sets.md#using-apis-from-later-requirement-sets)」を参照してください。
>
> <sup>2</sup> リリース後の更新プログラムで追加されました。
>
> <sup>3</sup>新しい Mac UI (プレビュー) のサポートは、Outlook バージョン 16.38.506 から利用できます。 詳細については、「[新しい Mac UI での Outlook のアドインのサポート](../outlook/compare-outlook-add-in-support-in-outlook-for-mac.md#add-in-support-in-outlook-on-new-mac-ui-preview)」セクションを参照してください。

> [!IMPORTANT]
> 要件セットのクライアント サポートは、Exchange サーバー サポートの制限を受ける場合があります。 Exchange サーバーおよび Outlook クライアントによってサポートされている要件セットの範囲の詳細については、「[Outlook JavaScript APIの要件セット](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)」を参照してください。

<br/>

## <a name="word"></a>Word

<table style="width:80%">
  <tr>
    <th>プラットフォーム</th>
    <th>拡張点</th>
    <th>API 要件セット</th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></th>
  </tr>
  <tr>
    <td>Office on the web</td>
    <td>
      - 作業ウィンドウ<br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets">IdentityAPI 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">設定</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Windows での Office<br>(Microsoft 365 サブスクリプションに接続)</td>
    <td>
      - 作業ウィンドウ<br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets">IdentityAPI 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/open-browser-window-api-requirement-sets">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">設定</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Windows 版 Office 2019<br>(1 回限りの購入)</td>
    <td>
      - 作業ウィンドウ<br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">設定</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Windows 版 Office 2016<br>(1 回限りの購入)</td>
    <td>- 作業ウィンドウ</td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*<br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">設定</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Windows 版 Office 2013<br>(1 回限りの購入)</td>
    <td>- 作業ウィンドウ</td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*<br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">設定</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Office on iPad<br>(Microsoft 365 サブスクリプションに接続)</td>
    <td>- 作業ウィンドウ</td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/open-browser-window-api-requirement-sets">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">設定</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Office on Mac<br>(Microsoft 365 サブスクリプションに接続)</td>
    <td>
      - 作業ウィンドウ<br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets">IdentityAPI 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/open-browser-window-api-requirement-sets">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">設定</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Mac 上の Office 2019<br>(1 回限りの購入)</td>
    <td>
      - 作業ウィンドウ<br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">設定</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Mac 上の Office 2016<br>(1 回限りの購入)</td>
    <td>- 作業ウィンドウ</td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*<br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">設定</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </td>
  </tr>
</table>

*&ast; - リリース後の更新プログラムで追加されました。*

<br/>

## <a name="powerpoint"></a>PowerPoint

<table style="width:80%">
  <tr>
    <th>プラットフォーム</th>
    <th>拡張点</th>
    <th>API 要件セット</th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></th>
  </tr>
  <tr>
    <td>Office on the web</td>
    <td>
      - コンテンツ<br>
      - 作業ウィンドウ<br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-1-1-requirement-set">PowerPointApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-1-2-requirement-set">PowerPointApi 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets">IdentityAPI 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Windows での Office<br>(Microsoft 365 サブスクリプションに接続)</td>
    <td>
      - コンテンツ<br>
      - 作業ウィンドウ<br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-1-1-requirement-set">PowerPointApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-1-2-requirement-set">PowerPointApi 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets">IdentityAPI 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/open-browser-window-api-requirement-sets">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Windows 版 Office 2019<br>(1 回限りの購入)</td>
    <td>
      - コンテンツ<br>
      - 作業ウィンドウ<br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Windows 版 Office 2016<br>(1 回限りの購入)</td>
    <td>
      - コンテンツ<br>
      - 作業ウィンドウ </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*<br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Windows 版 Office 2013<br>(1 回限りの購入)</td>
    <td>
      - コンテンツ<br>
      - 作業ウィンドウ </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*<br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office on iPad<br>(Microsoft 365 サブスクリプションに接続)</td>
    <td>
      - コンテンツ<br>
      - 作業ウィンドウ </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-1-1-requirement-set">PowerPointApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-1-2-requirement-set">PowerPointApi 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/open-browser-window-api-requirement-sets">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office on Mac<br>(Microsoft 365 サブスクリプションに接続)</td>
    <td>
      - コンテンツ<br>
      - 作業ウィンドウ<br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-1-1-requirement-set">PowerPointApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-1-2-requirement-set">PowerPointApi 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets">IdentityAPI 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/open-browser-window-api-requirement-sets">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Mac 上の Office 2019<br>(1 回限りの購入)</td>
    <td>
      - コンテンツ<br>
      - 作業ウィンドウ<br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Mac 上の Office 2016<br>(1 回限りの購入)</td>
    <td>
      - コンテンツ<br>
      - 作業ウィンドウ </td>
    <td>
       - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*<br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </td>
  </tr>
</table>

*&ast; - リリース後の更新プログラムで追加されました。*

<br/>

## <a name="onenote"></a>OneNote

<table style="width:80%">
  <tr>
    <th>プラットフォーム</th>
    <th>拡張点</th>
    <th>API 要件セット</th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></th>
  </tr>
  <tr>
    <td>Office on the web</td>
    <td>
      - コンテンツ<br>
      - 作業ウィンドウ<br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </td>
  </tr>
</table>

<br/>

## <a name="project"></a>Project

<table style="width:80%">
  <tr>
    <th>プラットフォーム</th>
    <th>拡張点</th>
    <th>API 要件セット</th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></th>
  </tr>
  <tr>
    <td>Windows 版 Office 2019<br>(1 回限りの購入)</td>
    <td>- 作業ウィンドウ</td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Windows 版 Office 2016<br>(1 回限りの購入)</td>
    <td>- 作業ウィンドウ</td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Windows 版 Office 2013<br>(1 回限りの購入)</td>
    <td>- 作業ウィンドウ</td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </td>
  </tr>
</table>

<br/>

## <a name="see-also"></a>関連項目

- [Office アドイン プラットフォームの概要](office-add-ins.md)
- [Office のバージョンと要件セット](../develop/office-versions-and-requirement-sets.md)
- [共通 API の要件セット](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [アドイン コマンドの要件セット](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [API リファレンス ドキュメント](../reference/javascript-api-for-office.md)
- [Microsoft 365 Apps の更新履歴](/officeupdates/update-history-office365-proplus-by-date)
- [Office 2016および2019の更新履歴（クリックして実行）](/officeupdates/update-history-office-2019)
- [Office 2013 の更新履歴 （クリックして実行）](/officeupdates/update-history-office-2013)
- [Office 2010、2013、および2016の更新履歴（MSI）](/officeupdates/office-updates-msi)
- [Outlook 2010、2013、および2016の更新履歴（MSI）](/officeupdates/outlook-updates-msi)
- [Office for Mac の更新履歴](/officeupdates/update-history-office-for-mac)
- [Office アドインを開発する](../develop/develop-overview.md)
