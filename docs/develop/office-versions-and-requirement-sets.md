---
title: Office のバージョンと要件セット
description: JavaScript API を使用してサポートされる Office.js プラットフォーム。
ms.date: 09/14/2022
ms.localizationpriority: high
ms.openlocfilehash: 669977f87974a1ec5519ddbbe3d38c5a290ec84f
ms.sourcegitcommit: cff5d3450f0c02814c1436f94cd1fc1537094051
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/30/2022
ms.locfileid: "68234908"
---
# <a name="office-versions-and-requirement-sets"></a>Office のバージョンと要件セット

Office にはプラットフォームやバージョンが異なるものが数多くあり、それらは Office JavaScript API (Office.js) に含まれる API をすべてサポートしているわけではありません。 Windows 上の Office 2013 は、Office アドインをサポートしていた最も古いバージョンの Office でした。ユーザーがインストールした Office のバージョンを常に制御できるわけではありません。 このような状況に対処するために、Office アプリケーションが Office アドインで必要な機能をサポートしているかどうかを判断するのに役立つ要件セットと呼ばれるシステムが用意されています。

> [!NOTE]
>
> - Office は、Windows、ブラウザー、Mac、iPad などの複数のプラットフォームで実行されます。
> - Office アプリケーションの例として、Excel、Word、PowerPoint、Outlook、OneNote などの Office 製品があります。
> - Office は、Microsoft 365 サブスクリプションまたは永続ライセンスで利用できます。 永続バージョンは、ボリューム ライセンス契約またはリテール版で入手できます。
> - 要件セットは、API メンバーの名前付きグループ (たとえば、API `WordApi 1.3`メンバーなど`ExcelApi 1.5`) です。

## <a name="how-to-check-your-office-version"></a>Office のバージョンを確認する方法

使用している Office のバージョンを特定するには、Office アプリケーション内で **[ファイル]** メニューを選択し、**[アカウント]** を選択します。 Office のバージョンが [ **Product Information** ] セクションに表示されます。 たとえば、次のスクリーンショットは、Office バージョン 1802 (ビルド 9026.1000) を示しています。

![Office のバージョン確認。](../images/office-version.png)

> [!NOTE]
> Office のバージョンがこれと異なる場合は、「自分 [が持っている Outlook のバージョン](https://support.microsoft.com/office/b3a9568c-edb5-42b9-9825-d48d82b2257c) 」または「 [Office について: 使用している Office のバージョン](https://support.microsoft.com/topic/932788b8-a3ce-44bf-bb09-e334518b8b19) 」を参照して、バージョンのこの情報を取得する方法を理解してください。

## <a name="office-requirement-sets-availability"></a>Office 要件セットの可用性

Office アドインでは、API 要件セットを使用して、Office アプリケーションが使用する必要がある API メンバーをサポートしているかどうかを判断できます。 要件セットのサポートは、Office アプリケーションと Office アプリケーションのバージョンによって異なります (前のセクション「 [Office バージョンを確認する方法](#how-to-check-your-office-version)」を参照)。

Some Office applications have their own API requirement sets. For example, the first requirement set for the Excel API was `ExcelApi 1.1` and the first requirement set for the Word API was `WordApi 1.1`. Since then, multiple new ExcelApi requirement sets and WordApi requirement sets have been added to provide additional API functionality.

さらに、アドイン コマンド (リボン機能拡張) やダイアログ ボックスを起動する機能 (ダイアログ API) など、他の機能が共通 API に追加されました。 アドイン コマンドと Dialog API 要件セットは、さまざまな Office アプリケーションが共通する API セットの例です。

An add-in can only use APIs in requirement sets that are supported by the version of Office application where the add-in is running. To know exactly which requirement sets are available for a specific Office application version, refer to the following application-specific requirement set articles.

- [Excel JavaScript API 要件セット](/javascript/api/requirement-sets/excel/excel-api-requirement-sets) (ExcelApi)
- [OneNote JavaScript API 要件セット](/javascript/api/requirement-sets/onenote/onenote-api-requirement-sets) (OneNoteApi)
- [Outlook JavaScript API 要件セット](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets) (メールボックス)
- [PowerPoint JavaScript API 要件セット](/javascript/api/requirement-sets/powerpoint/powerpoint-api-requirement-sets) (PowerPointApi)
- [Word JavaScript API 要件セット](/javascript/api/requirement-sets/word/word-api-requirement-sets) (WordApi)

一部の要件セットには、複数の Office アプリケーションで使用できる API が含まれています。 これらの要件セットの詳細については、次の記事を参照してください。

- [Office の共通要件セット](/javascript/api/requirement-sets/common/office-add-in-requirement-sets)
- [アドイン コマンドの要件セット](/javascript/api/requirement-sets/common/add-in-commands-requirement-sets)
- [ダイアログ API の要件セット](/javascript/api/requirement-sets/common/dialog-api-requirement-sets)
- [ダイアログ配信元の要件セット](/javascript/api/requirement-sets/common/dialog-origin-requirement-sets)
- [Identity API の要件セット](/javascript/api/requirement-sets/common/identity-api-requirement-sets)
- [画像強制型変換要件セット](/javascript/api/requirement-sets/common/image-coercion-requirement-sets)
- [キーボード ショートカットの要件セット](/javascript/api/requirement-sets/common/keyboard-shortcuts-requirement-sets)
- [ブラウザー ウィンドウの要件セットを開く](/javascript/api/requirement-sets/common/open-browser-window-api-requirement-sets)
- [リボン API の要件セット](/javascript/api/requirement-sets/common/ribbon-api-requirement-sets)
- [共有ランタイム要件のセット](/javascript/api/requirement-sets/common/shared-runtime-requirement-sets)

The version number of a requirement set, such as the "1.1" in `ExcelApi 1.1`, is relative to the Office application. The version number of a given requirement set (e.g., `ExcelApi 1.1`) does not correspond to the version number of Office.js or to requirement sets for other Office applications (e.g., Word, Outlook, etc.).  Requirement sets for the different Office applications are released at different rates. For example, `ExcelApi 1.5` was released before the `WordApi 1.3` requirement set.

The Office JavaScript API library (Office.js) includes all requirement sets that are currently available. While there is such a thing as requirement sets `ExcelApi 1.3` and `WordApi 1.3`, there is no `Office.js 1.3` requirement set. The latest release of Office.js is maintained as a single Office endpoint delivered via the content delivery network (CDN). For more details around the Office.js CDN, including how versioning and backward compatibility is handled, see [Understanding the Office JavaScript API](../develop/understanding-the-javascript-api-for-office.md).

## <a name="specify-office-applications-and-requirement-sets"></a>Office アプリケーションと要件セットを指定する

There are various ways to specify which Office applications and requirement sets are required by an add-in.  For detailed information, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md)

## <a name="see-also"></a>関連項目

- [Office アプリケーションと API の要件を指定する](../develop/specify-office-hosts-and-api-requirements.md)
- [Office の最新バージョンをインストールする](../develop/install-latest-office-version.md)
- [Microsoft 365 Apps 用更新プログラム チャネルの概要](/deployoffice/overview-of-update-channels-for-office-365-proplus)
- [Microsoft 365 と Microsoft Teams による生産性の再構築](https://products.office.com/compare-all-microsoft-office-products?tab=2)
