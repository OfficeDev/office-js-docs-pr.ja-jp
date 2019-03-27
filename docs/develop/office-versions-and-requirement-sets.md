---
title: Office のバージョンと要件セット
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 1af4840e965d7043505aedc8330cebacacc4d203
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30872208"
---
# <a name="office-versions-and-requirement-sets"></a><span data-ttu-id="38444-102">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="38444-102">Office versions and requirement sets</span></span>

<span data-ttu-id="38444-p101">Office にはプラットフォームやバージョンが異なるものが数多くあり、それらは Office JavaScript API (Office.js) に含まれる API をすべてサポートしているわけではありません。 ユーザーがインストールしている Office のバージョンを制御できない場合があります。  このような状況に対処するため、Office アドインで必要な機能を Office ホストがサポートしているかどうかを判別するのに役立つ要件セットと呼ばれるシステムが用意されています。</span><span class="sxs-lookup"><span data-stu-id="38444-p101">There are many versions of Office on several platforms, and they don't all support every API in Office JavaScript API (Office.js). You may not always have control over the version of Office your users have installed.  To handle this situation, we provide a system called requirement sets to help you determine whether an Office host supports the capabilities you need in your Office Add-in.</span></span> 

> [!NOTE]
> - <span data-ttu-id="38444-106">Office は、Office for Windows、Office Online、Office for Mac、Office for iPad を含む複数のプラットフォームで実行できます。</span><span class="sxs-lookup"><span data-stu-id="38444-106">Office runs across multiple platforms, including Office for Windows, Office Online, Office for the Mac, and Office for the iPad.</span></span>
> - <span data-ttu-id="38444-107">Office ホストの例は、Excel、Word、PowerPoint、Outlook、OneNote などの Office 製品です。</span><span class="sxs-lookup"><span data-stu-id="38444-107">Examples of Office hosts are Office Products: Excel, Word, PowerPoint, Outlook, OneNote, and so forth.</span></span>  
> - <span data-ttu-id="38444-108">要件セットとは、`ExcelApi 1.5` や `WordApi 1.3` などの、API メンバーの名前付きグループです。</span><span class="sxs-lookup"><span data-stu-id="38444-108">A requirement set is a named group of API members e.g., `ExcelApi 1.5`, `WordApi 1.3`, and so on.</span></span>  


## <a name="how-to-check-your-office-version"></a><span data-ttu-id="38444-109">Office のバージョンを確認する方法</span><span class="sxs-lookup"><span data-stu-id="38444-109">How to check your Office version</span></span>

<span data-ttu-id="38444-p102">使用している Office のバージョンを特定するには、Office アプリケーション内で **[ファイル]** メニューを選択し、**[アカウント]** を選択します。 Office のバージョンは **[製品情報]** セクションに表示されます。 たとえば、次のスクリーン ショットは、Office のバージョンが 1802 (ビルド 9026.1000) であることを示しています。</span><span class="sxs-lookup"><span data-stu-id="38444-p102">To identify the Office version that you're using, from within an Office application, select the **File** menu, and then choose **Account**. The version of Office will appear in the **Product Information** section. For example, the following screenshot indicates Office Version 1802 (Build 9026.1000):</span></span>

![Office のバージョン確認](../images/office-version-number-ui.jpg)


## <a name="office-requirement-sets-availability"></a><span data-ttu-id="38444-114">Office 要件セットの可用性</span><span class="sxs-lookup"><span data-stu-id="38444-114">Office requirement sets availability</span></span>

<span data-ttu-id="38444-p103">Office アドインは API 要件セットを使用して、使用する必要のある API メンバーを Office ホストがサポートしているかどうかを判別できます。 要件セットのサポートは、Office ホストと Office ホストのバージョンによって異なります (前のセクションを参照してください)。</span><span class="sxs-lookup"><span data-stu-id="38444-p103">Office Add-ins can use API requirement sets to determine whether the Office host supports the API members that it need to use. Requirement set support varies by Office host and the Office host version (see previous section).</span></span>

<span data-ttu-id="38444-p104">一部の Office ホストには独自の API 要件セットがあります。 たとえば、Excel API の最初の要件セットは `ExcelApi 1.1` で、Word API の最初の要件セットは `WordApi 1.1` でした。 それ以降、追加の API 機能を提供するため、複数の新しい ExcelApi 要件セットと WordApi 要件セットが追加されています。</span><span class="sxs-lookup"><span data-stu-id="38444-p104">Some Office hosts have their own API requirement sets. For example, the first requirement set for the Excel API was `ExcelApi 1.1` and the first requirement set for the Word API was `WordApi 1.1`. Since then, multiple new ExcelApi requirement sets and WordApi requirement sets have been added to provide additional API functionality.</span></span>

<span data-ttu-id="38444-120">さらに、アドイン コマンド (リボン機能拡張) やダイアログ ボックスを起動する機能 (ダイアログ API) など、他の機能が共通 API に追加されました。</span><span class="sxs-lookup"><span data-stu-id="38444-120">In addition, other functionality such as add-in commands (ribbon extensibility) and the ability to launch dialog boxes (Dialog API) were added to the Common API.</span></span> <span data-ttu-id="38444-121">アドイン コマンドやダイアログ API の要件セットは、さまざまな Office ホストで共有されている API セットの例です。</span><span class="sxs-lookup"><span data-stu-id="38444-121">Add-in commands and Dialog API requirement sets are examples of API sets that the various Office hosts share in common.</span></span>

<span data-ttu-id="38444-p106">アドインは、そのアドインが動作している Office ホストのバージョンでサポートしている要件セットにある API のみを使用できます。 特定の Office ホストのバージョンで使用できる要件セットを正確に確認するには、ホスト固有の要件セットに関する次の記事を参照してください。</span><span class="sxs-lookup"><span data-stu-id="38444-p106">An add-in can only use APIs in requirement sets that are supported by the version of Office host where the add-in is running. To know exactly which requirement sets are available for a specific Office host version, refer to the following host-specific requirement set articles:</span></span>

- <span data-ttu-id="38444-124">[Excel JavaScript API 要件セット](/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets) (ExcelApi)</span><span class="sxs-lookup"><span data-stu-id="38444-124">[Excel JavaScript API requirement sets](/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets) (ExcelApi)</span></span>
- <span data-ttu-id="38444-125">[Word JavaScript API 要件セット](/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets) (WordApi)</span><span class="sxs-lookup"><span data-stu-id="38444-125">[Word JavaScript API requirement sets](/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets) (WordApi)</span></span>
- <span data-ttu-id="38444-126">[OneNote JavaScript API 要件セット](/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets) (OneNoteApi)</span><span class="sxs-lookup"><span data-stu-id="38444-126">[OneNote JavaScript API requirement sets](/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets) (OneNoteApi)</span></span>
- <span data-ttu-id="38444-127">[Outlook API 要件セットについて](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) (MailBox)</span><span class="sxs-lookup"><span data-stu-id="38444-127">[Understanding Outlook API requirement sets](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) (MailBox)</span></span>

<span data-ttu-id="38444-p107">一部の要件セットには、どの Office ホストでも使用できる API が含まれています。 これらの要件のセットの詳細については、次の記事を参照してください。</span><span class="sxs-lookup"><span data-stu-id="38444-p107">Some requirement sets contain APIs that can be used by any Office host. For information about these requirement sets, refer to the following articles:</span></span>

- [<span data-ttu-id="38444-130">Office の共通要件セット</span><span class="sxs-lookup"><span data-stu-id="38444-130">Office common requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="38444-131">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="38444-131">Add-in commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="38444-132">ダイアログ API の要件セット</span><span class="sxs-lookup"><span data-stu-id="38444-132">Dialog API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets)
- [<span data-ttu-id="38444-133">Identity API の要件セット</span><span class="sxs-lookup"><span data-stu-id="38444-133">Identity API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets)

<span data-ttu-id="38444-p108">`ExcelApi 1.1` の "1.1" など、要件セットのバージョン番号は Office ホストを基準にしています。 特定の要件セットのバージョン番号 (例: `ExcelApi 1.1`) は、Office.js のバージョン番号には対応しておらず、他の Office ホスト (Word、Outlook など) の要件セットにも対応していません。  Office ホストの要件セットがリリースされる早さや時期は、ホストによって異なります。 たとえば、`ExcelApi 1.5` の方が `WordApi 1.3` 要件セットより前にリリースされました。</span><span class="sxs-lookup"><span data-stu-id="38444-p108">The version number of a requirement set, such as the "1.1" in `ExcelApi 1.1`, is relative to the Office host. The version number of a given requirement set (e.g., `ExcelApi 1.1`) does not correspond to the version number of Office.js or to requirement sets for other Office hosts (e.g., Word, Outlook, etc.).  Requirement sets for the different Office hosts are released at different speeds and times. For example, `ExcelApi 1.5` was released before the `WordApi 1.3` requirement set.</span></span>

<span data-ttu-id="38444-138">JavaScript API for Office ライブラリ (Office.js) には、現在利用可能なすべての要件セットが含まれています。</span><span class="sxs-lookup"><span data-stu-id="38444-138">The JavaScript API for Office library (Office.js) includes all requirement sets that are currently available.</span></span> <span data-ttu-id="38444-139">`ExcelApi 1.3` や `WordApi 1.3` のような要件セットは存在しますが、`Office.js 1.3` のような要件セットは存在しません。</span><span class="sxs-lookup"><span data-stu-id="38444-139">While there is such a thing as requirement sets `ExcelApi 1.3` and `WordApi 1.3`, there is no `Office.js 1.3` requirement set.</span></span> <span data-ttu-id="38444-140">Office.js の最新リリースは、コンテンツ配信ネットワーク (CDN) 経由で配信される単一の Office エンドポイントとして維持されます。</span><span class="sxs-lookup"><span data-stu-id="38444-140">The latest release of Office.js is maintained as a single Office endpoint delivered via the content delivery network (CDN).</span></span> <span data-ttu-id="38444-141">バージョン管理や下位互換性の処理方法など、Office.js CDN に関する詳細については、「[JavaScript API for Office について](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="38444-141">For more details around the Office.js CDN, including how versioning and backward compatibility is handled, see [Understanding the JavaScript API for Office](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span></span>

## <a name="specify-office-hosts-and-requirement-sets"></a><span data-ttu-id="38444-142">Office ホストと要件セットを指定する</span><span class="sxs-lookup"><span data-stu-id="38444-142">Specify Office hosts and requirement sets</span></span>

<span data-ttu-id="38444-p110">アドインに必要となる Office ホストと要件セットは、さまざまな方法で指定できます。  詳細については、「[Office のホストと API の要件を指定する](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="38444-p110">There are various ways to specify which Office hosts and requirement sets are required by an add-in.  For detailed information, see [Specify Office hosts and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)</span></span>


## <a name="see-also"></a><span data-ttu-id="38444-145">関連項目</span><span class="sxs-lookup"><span data-stu-id="38444-145">See also</span></span>

- [<span data-ttu-id="38444-146">Office のホストと API の要件を指定する</span><span class="sxs-lookup"><span data-stu-id="38444-146">Specify Office hosts and API requirements</span></span>](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="38444-147">Office の最新バージョンをインストールする</span><span class="sxs-lookup"><span data-stu-id="38444-147">Install the latest version of Office</span></span>](/office/dev/add-ins/develop/install-latest-office-version)
- [<span data-ttu-id="38444-148">Office 365 ProPlus 更新プログラムのチャネルの概要</span><span class="sxs-lookup"><span data-stu-id="38444-148">Overview of update channels for Office 365 ProPlus</span></span>](/deployoffice/overview-of-update-channels-for-office-365-proplus)
- [<span data-ttu-id="38444-149">Office 365 で Office を最大限に活用する</span><span class="sxs-lookup"><span data-stu-id="38444-149">Get the most from Office with Office 365</span></span>](https://products.office.com/compare-all-microsoft-office-products?tab=2)
