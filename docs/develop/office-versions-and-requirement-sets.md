---
title: Office のバージョンと要件セット
description: JavaScript API を使用してサポートされる Office.js プラットフォーム
ms.date: 07/07/2020
localization_priority: Priority
ms.openlocfilehash: 02f3d91256ea05e526ebe2e4e4090b1908d7292a
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093582"
---
# <a name="office-versions-and-requirement-sets"></a>Office のバージョンと要件セット

There are many versions of Office on several platforms, and they don't all support every API in Office JavaScript API (Office.js). You may not always have control over the version of Office your users have installed.  To handle this situation, we provide a system called requirement sets to help you determine whether an Office host supports the capabilities you need in your Office Add-in. 

> [!NOTE]
> - Office は、Windows、ブラウザー、Mac、iPad などの複数のプラットフォームで実行されます。
> - Office ホストの例は、Excel、Word、PowerPoint、Outlook、OneNote などの Office 製品です。  
> - 要件セットとは、`ExcelApi 1.5` や `WordApi 1.3` などの、API メンバーの名前付きグループです。  

## <a name="how-to-check-your-office-version"></a>Office のバージョンを確認する方法

To identify the Office version that you're using, from within an Office application, select the **File** menu, and then choose **Account**. The version of Office will appear in the **Product Information** section. For example, the following screenshot indicates Office Version 1802 (Build 9026.1000):

![Office のバージョン確認](../images/office-version.png)

## <a name="office-requirement-sets-availability"></a>Office 要件セットの可用性

Office Add-ins can use API requirement sets to determine whether the Office host supports the API members that it need to use. Requirement set support varies by Office host and the Office host version (see previous section).

Some Office hosts have their own API requirement sets. For example, the first requirement set for the Excel API was `ExcelApi 1.1` and the first requirement set for the Word API was `WordApi 1.1`. Since then, multiple new ExcelApi requirement sets and WordApi requirement sets have been added to provide additional API functionality.

さらに、アドイン コマンド (リボン機能拡張) やダイアログ ボックスを起動する機能 (ダイアログ API) など、他の機能が共通 API に追加されました。 アドイン コマンドやダイアログ API の要件セットは、さまざまな Office ホストで共有されている API セットの例です。

An add-in can only use APIs in requirement sets that are supported by the version of Office host where the add-in is running. To know exactly which requirement sets are available for a specific Office host version, refer to the following host-specific requirement set articles:

- [Excel JavaScript API 要件セット](../reference/requirement-sets/excel-api-requirement-sets.md) (ExcelApi)
- [Word JavaScript API 要件セット](../reference/requirement-sets/word-api-requirement-sets.md) (WordApi)
- [OneNote JavaScript API 要件セット](../reference/requirement-sets/onenote-api-requirement-sets.md) (OneNoteApi)
- [PowerPoint JavaScript API 要件セット](../reference/requirement-sets/powerpoint-api-requirement-sets.md) (PowerPointApi)
- [Outlook API 要件セットについて](../reference/requirement-sets/outlook-api-requirement-sets.md) (Mailbox)

Some requirement sets contain APIs that can be used by any Office host. For information about these requirement sets, refer to the following articles:

- [Office の共通要件セット](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [アドイン コマンドの要件セット](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [ダイアログ API の要件セット](../reference/requirement-sets/dialog-api-requirement-sets.md)
- [Identity API の要件セット](../reference/requirement-sets/identity-api-requirement-sets.md)

The version number of a requirement set, such as the "1.1" in `ExcelApi 1.1`, is relative to the Office host. The version number of a given requirement set (e.g., `ExcelApi 1.1`) does not correspond to the version number of Office.js or to requirement sets for other Office hosts (e.g., Word, Outlook, etc.).  Requirement sets for the different Office hosts are released at different speeds and times. For example, `ExcelApi 1.5` was released before the `WordApi 1.3` requirement set.

Office JavaScript API ライブラリ (Office.js) には、現在利用可能なすべての要件セットが含まれています。 `ExcelApi 1.3` や `WordApi 1.3` のような要件セットは存在しますが、`Office.js 1.3` のような要件セットは存在しません。 Office.js の最新リリースは、コンテンツ配信ネットワーク (CDN) 経由で配信される単一の Office エンドポイントとして維持されます。 バージョン管理や下位互換性の処理方法など、Office.js CDN に関する詳細については、「[Office JavaScript API について](../develop/understanding-the-javascript-api-for-office.md)」を参照してください。

## <a name="specify-office-hosts-and-requirement-sets"></a>Office ホストと要件セットを指定する

There are various ways to specify which Office hosts and requirement sets are required by an add-in.  For detailed information, see [Specify Office hosts and API requirements](../develop/specify-office-hosts-and-api-requirements.md)

## <a name="see-also"></a>関連項目

- [Office のホストと API の要件を指定する](../develop/specify-office-hosts-and-api-requirements.md)
- [Office の最新バージョンをインストールする](../develop/install-latest-office-version.md)
- [Microsoft 365 Apps 用更新プログラム チャネルの概要](/deployoffice/overview-of-update-channels-for-office-365-proplus)
- [Office 365 で Office を最大限に活用する](https://products.office.com/compare-all-microsoft-office-products?tab=2)
