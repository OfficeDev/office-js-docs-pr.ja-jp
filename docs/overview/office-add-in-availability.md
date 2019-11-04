---
title: Office アドインを使用できるホストおよびプラットフォーム
description: Excel、OneNote、Outlook、PowerPoint、Project、Word のサポートされる要件セット。
ms.date: 10/30/2019
localization_priority: Priority
ms.openlocfilehash: 3621236ea86410d70d17655450e1f6d32a212823
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/31/2019
ms.locfileid: "37901949"
---
# <a name="office-add-in-host-and-platform-availability"></a>Office アドインを使用できるホストおよびプラットフォーム

期待どおりの動作をするうえで、Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、および各 Office アプリケーションで現在サポートされている共通 API が含まれています。

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
    <td> - 作業ウィンドウ<br>
        - コンテンツ<br>
        - カスタム関数<br>
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
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>
        - BindingEvents<br>
        - CompressedFile<br>
        - DocumentEvents<br>
        - File<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - Selection<br>
        - Settings<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Windows での Office<br>(Office 365 サブスクリプションに接続済み)</td>
    <td> - 作業ウィンドウ<br>
        - コンテンツ<br>
        - カスタム関数<br>
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
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></td>
    <td>
        - BindingEvents<br>
        - CompressedFile<br>
        - DocumentEvents<br>
        - File<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - Selection<br>
        - Settings<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Windows 版 Office 2019<br>(1 回限りの購入)</td>
    <td>- 作業ウィンドウ<br>
        - コンテンツ<br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td>- BindingEvents<br>
        - CompressedFile<br>
        - DocumentEvents<br>
        - File<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - Selection<br>
        - Settings<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Windows 版 Office 2016<br>(1 回限りの購入)</td>
    <td>- 作業ウィンドウ<br>
        - コンテンツ</td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*<br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td>- BindingEvents<br>
        - CompressedFile<br>
        - DocumentEvents<br>
        - File<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - Selection<br>
        - Settings<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Windows 版 Office 2013<br>(1 回限りの購入)</td>
    <td>
        - 作業ウィンドウ<br>
        - コンテンツ</td>
    <td>  - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*<br>
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td>
        - BindingEvents<br>
        - CompressedFile<br>
        - DocumentEvents<br>
        - File<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - Selection<br>
        - Settings<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>iPad 上の Office<br>(Office 365 サブスクリプションに接続済み)</td>
    <td>- 作業ウィンドウ<br>
        - コンテンツ</td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td>- BindingEvents<br>
        - DocumentEvents<br>
        - File<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - Selection<br>
        - Settings<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Mac 上の Office<br>(Office 365 サブスクリプションに接続済み)</td>
    <td>- 作業ウィンドウ<br>
        - コンテンツ<br>
        - カスタム関数<br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></td>
    <td>- BindingEvents<br>
        - CompressedFile<br>
        - DocumentEvents<br>
        - File<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - PdfFile<br>
        - Selection<br>
        - Settings<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Mac 上の Office 2019<br>(1 回限りの購入)</td>
    <td>- 作業ウィンドウ<br>
        - コンテンツ<br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td>- BindingEvents<br>
        - CompressedFile<br>
        - DocumentEvents<br>
        - File<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - PdfFile<br>
        - Selection<br>
        - Settings<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Mac 上の Office 2016<br>(1 回限りの購入)</td>
    <td>- 作業ウィンドウ<br>
        - コンテンツ</td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*<br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td>- BindingEvents<br>
        - CompressedFile<br>
        - DocumentEvents<br>
        - File<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - PdfFile<br>
        - Selection<br>
        - Settings<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
</table>

*&ast; - リリース後の更新プログラムで追加されました。*

## <a name="custom-functions"></a>カスタム関数

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
        - カスタム関数</td>
    <td>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></td>
    <td>
    </td>
  </tr>
  <tr>
    <td>Windows での Office<br>(Office 365 サブスクリプションに接続済み)</td>
    <td>
        - カスタム関数</td>
    <td>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></td>
    <td>
    </td>
  </tr>
  <tr>
    <td>Office for Mac<br>(Office 365 に接続された)</td>
    <td>
        - カスタム関数</td>
    <td>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></td>
    <td>
    </td>
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
    <td> - メールの読み取り<br>
      - メールの作成<br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></td>
    <td>利用不可</td>
  </tr>
  <tr>
    <td>Office on the web<br>(クラシック)</td>
    <td> - メールの読み取り<br>
      - メールの作成<br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></td>
    <td>使用不可</td>
  </tr>
  <tr>
    <td>Windows での Office<br>(Office 365 サブスクリプションに接続済み)</td>
    <td> - メールの読み取り<br>
      - メールの作成<br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a><br>
      - モジュール</td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></td>
    <td>利用不可</td>
  </tr>
  <tr>
    <td>Windows 版 Office 2019<br>(1 回限りの購入)</td>
    <td> - メールの読み取り<br>
      - メールの作成<br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a><br>
      - モジュール</td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></td>
    <td>使用不可</td>
  </tr>
  <tr>
    <td>Windows 版 Office 2016<br>(1 回限りの購入)</td>
    <td> - メールの読み取り<br>
      - メールの作成<br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a><br>
      - モジュール</td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</td>
    <td>使用不可</td>
  </tr>
  <tr>
    <td>Windows 版 Office 2013<br>(1 回限りの購入)</td>
    <td> - メールの読み取り<br>
      - メールの作成</td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*<br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</td>
    <td>使用不可</td>
  </tr>
  <tr>
    <td>iOS 上の Office<br>(Office 365 サブスクリプションに接続済み)</td>
    <td> - メールの読み取り<br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></td>
    <td>使用不可</td>
  </tr>
  <tr>
    <td>Mac 上の Office<br>(Office 365 サブスクリプションに接続済み)</td>
    <td> - メールの読み取り<br>
      - メールの作成<br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></td>
    <td>利用不可</td>
  </tr>
  <tr>
    <td>Mac 上の Office 2019<br>(1 回限りの購入)</td>
    <td> - メールの読み取り<br>
      - メールの作成<br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></td>
    <td>使用不可</td>
  </tr>
  <tr>
    <td>Mac 上の Office 2016<br>(1 回限りの購入)</td>
    <td> - メールの読み取り<br>
      - メールの作成<br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></td>
    <td>使用不可</td>
  </tr>
  <tr>
    <td>Android 上の Office<br>(Office 365 サブスクリプションに接続済み)</td>
    <td> - メールの読み取り<br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></td>
    <td>利用不可</td>
  </tr>
</table>

*&ast; - リリース後の更新プログラムで追加されました。*

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
    <td> - 作業ウィンドウ<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></td>
    <td> - BindingEvents<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile</td>
  </tr>
  <tr>
    <td>Windows での Office<br>(Office 365 サブスクリプションに接続済み)</td>
    <td> - 作業ウィンドウ<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a><br>
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile </td>
  </tr>
  <tr>
    <td>Windows 版 Office 2019<br>(1 回限りの購入)</td>
    <td> - 作業ウィンドウ<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile </td>
  </tr>
  <tr>
    <td>Windows 版 Office 2016<br>(1 回限りの購入)</td>
    <td> - 作業ウィンドウ</td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*<br>
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile </td>
  </tr>
  <tr>
    <td>Windows 版 Office 2013<br>(1 回限りの購入)</td>
    <td> - 作業ウィンドウ</td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*<br>
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile</td>
  </tr>
  <tr>
    <td>iPad 上の Office<br>(Office 365 サブスクリプションに接続済み)</td>
    <td> - 作業ウィンドウ</td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
</td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile </td>
  </tr>
  <tr>
    <td>Mac 上の Office<br>(Office 365 サブスクリプションに接続済み)</td>
    <td> - 作業ウィンドウ<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></td>
</td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile </td>
  </tr>
  <tr>
    <td>Mac 上の Office 2019<br>(1 回限りの購入)</td>
    <td> - 作業ウィンドウ<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
</td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile </td>
  </tr>
  <tr>
    <td>Mac 上の Office 2016<br>(1 回限りの購入)</td>
    <td> - 作業ウィンドウ</td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*<br>
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile </td>
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
    <td> - コンテンツ<br>
         - 作業ウィンドウ<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Windows での Office<br>(Office 365 サブスクリプションに接続済み)</td>
    <td> - コンテンツ<br>
         - 作業ウィンドウ<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Windows 版 Office 2019<br>(1 回限りの購入)</td>
    <td> - コンテンツ<br>
         - 作業ウィンドウ<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Windows 版 Office 2016<br>(1 回限りの購入)</td>
    <td> - コンテンツ<br>
         - 作業ウィンドウ</td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Windows 版 Office 2013<br>(1 回限りの購入)</td>
    <td> - コンテンツ<br>
         - 作業ウィンドウ<br>
    </td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>iPad 上の Office<br>(Office 365 サブスクリプションに接続済み)</td>
    <td> - コンテンツ<br>
         - 作業ウィンドウ</td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Mac 上の Office<br>(Office 365 サブスクリプションに接続済み)</td>
    <td> - コンテンツ<br>
         - 作業ウィンドウ<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Mac 上の Office 2019<br>(1 回限りの購入)</td>
    <td> - コンテンツ<br>
         - 作業ウィンドウ<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Mac 上の Office 2016<br>(1 回限りの購入)</td>
    <td> - コンテンツ<br>
         - 作業ウィンドウ</td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion</td>
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
    <td> - コンテンツ<br>
         - 作業ウィンドウ<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td> - DocumentEvents<br>
         - HtmlCoercion<br>
         - Settings<br>
         - TextCoercion</td>
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
    <td> - 作業ウィンドウ</td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - Selection<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Windows 版 Office 2016<br>(1 回限りの購入)</td>
    <td> - 作業ウィンドウ</td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - Selection<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Windows 版 Office 2013<br>(1 回限りの購入)</td>
    <td> - 作業ウィンドウ</td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - Selection<br>
         - TextCoercion</td>
  </tr>
</table>

<br/>

## <a name="see-also"></a>関連項目

- [Office アドイン プラットフォームの概要](office-add-ins.md)
- [Office のバージョンと要件セット](../develop/office-versions-and-requirement-sets.md)
- [共通 API の要件セット](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [アドイン コマンドの要件セット](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [JavaScript API for Office リファレンス](../reference/javascript-api-for-office.md)
- [Office 365 ProPlus の更新履歴](/officeupdates/update-history-office365-proplus-by-date)
- [Office 2016および2019の更新履歴（クリックして実行）](/officeupdates/update-history-office-2019)
- [Office 2013 の更新履歴 （クリックして実行）](/officeupdates/update-history-office-2013)
- [Office 2010、2013、および2016の更新履歴（MSI）](/officeupdates/office-updates-msi)
- [Outlook 2010、2013、および2016の更新履歴（MSI）](/officeupdates/outlook-updates-msi)
- [Office for Mac の更新履歴](/officeupdates/update-history-office-for-mac)
