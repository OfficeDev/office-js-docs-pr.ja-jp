---
title: Office のバージョンと要件セット
description: JavaScript API を使用してサポートされる Office.js プラットフォーム。
ms.date: 02/09/2021
localization_priority: Priority
ms.openlocfilehash: 65db7bf6e8670e389cfaf5e557b365d960376569
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075993"
---
# <a name="office-versions-and-requirement-sets"></a><span data-ttu-id="abaf1-103">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="abaf1-103">Office versions and requirement sets</span></span>

<span data-ttu-id="abaf1-p101">Office にはプラットフォームやバージョンが異なるものが数多くあり、それらは Office JavaScript API (Office.js) に含まれる API をすべてサポートしているわけではありません。 ユーザーがインストールしている Office のバージョンを制御できない場合があります。このような状況に対処するため、Office アドインで必要な機能を Office アプリケーションがサポートしているかどうかを判別するのに役立つ要件セットと呼ばれるシステムが用意されています。</span><span class="sxs-lookup"><span data-stu-id="abaf1-p101">There are many versions of Office on several platforms, and they don't all support every API in Office JavaScript API (Office.js). You may not always have control over the version of Office your users have installed.  To handle this situation, we provide a system called requirement sets to help you determine whether an Office application supports the capabilities you need in your Office Add-in.</span></span> 

> [!NOTE]
> - <span data-ttu-id="abaf1-107">Office は、Windows、ブラウザー、Mac、iPad などの複数のプラットフォームで実行されます。</span><span class="sxs-lookup"><span data-stu-id="abaf1-107">Office runs across multiple platforms, including Windows, in a browser, Mac, and iPad.</span></span>
> - <span data-ttu-id="abaf1-108">Office アプリケーションの例は、Excel、Word、PowerPoint、Outlook、OneNote などの Office 製品です。</span><span class="sxs-lookup"><span data-stu-id="abaf1-108">Examples of Office applications are Office Products: Excel, Word, PowerPoint, Outlook, OneNote, and so forth.</span></span>  
> - <span data-ttu-id="abaf1-109">要件セットとは、`ExcelApi 1.5` や `WordApi 1.3` などの、API メンバーの名前付きグループです。</span><span class="sxs-lookup"><span data-stu-id="abaf1-109">A requirement set is a named group of API members e.g., `ExcelApi 1.5`, `WordApi 1.3`, and so on.</span></span>  

## <a name="how-to-check-your-office-version"></a><span data-ttu-id="abaf1-110">Office のバージョンを確認する方法</span><span class="sxs-lookup"><span data-stu-id="abaf1-110">How to check your Office version</span></span>

<span data-ttu-id="abaf1-p102">使用している Office のバージョンを特定するには、Office アプリケーション内で **[ファイル]** メニューを選択し、**[アカウント]** を選択します。 Office のバージョンは **[製品情報]** セクションに表示されます。 たとえば、次のスクリーン ショットは、Office のバージョンが 1802 (ビルド 9026.1000) であることを示しています。</span><span class="sxs-lookup"><span data-stu-id="abaf1-p102">To identify the Office version that you're using, from within an Office application, select the **File** menu, and then choose **Account**. The version of Office will appear in the **Product Information** section. For example, the following screenshot indicates Office Version 1802 (Build 9026.1000):</span></span>

![Office のバージョン確認。](../images/office-version.png)

## <a name="office-requirement-sets-availability"></a><span data-ttu-id="abaf1-115">Office 要件セットの可用性</span><span class="sxs-lookup"><span data-stu-id="abaf1-115">Office requirement sets availability</span></span>

<span data-ttu-id="abaf1-p103">Office アドインは API 要件セットを使用して、使用する必要のある API メンバーを Office アプリケーションがサポートしているかどうかを判別できます。 要件セットのサポートは、Office アプリケーションと Office アプリケーションのバージョンによって異なります (前のセクションを参照してください)。</span><span class="sxs-lookup"><span data-stu-id="abaf1-p103">Office Add-ins can use API requirement sets to determine whether the Office application supports the API members that it need to use. Requirement set support varies by Office application and the Office application version (see previous section).</span></span>

<span data-ttu-id="abaf1-p104">一部の Office アプリケーションには独自の API 要件セットがあります。 たとえば、Excel API の最初の要件セットは `ExcelApi 1.1` で、Word API の最初の要件セットは `WordApi 1.1` でした。 それ以降、追加の API 機能を提供するため、複数の新しい ExcelApi 要件セットと WordApi 要件セットが追加されています。</span><span class="sxs-lookup"><span data-stu-id="abaf1-p104">Some Office applications have their own API requirement sets. For example, the first requirement set for the Excel API was `ExcelApi 1.1` and the first requirement set for the Word API was `WordApi 1.1`. Since then, multiple new ExcelApi requirement sets and WordApi requirement sets have been added to provide additional API functionality.</span></span>

<span data-ttu-id="abaf1-121">さらに、アドイン コマンド (リボン機能拡張) やダイアログ ボックスを起動する機能 (ダイアログ API) など、他の機能が共通 API に追加されました。</span><span class="sxs-lookup"><span data-stu-id="abaf1-121">In addition, other functionality such as add-in commands (ribbon extensibility) and the ability to launch dialog boxes (Dialog API) were added to the Common API.</span></span> <span data-ttu-id="abaf1-122">アドイン コマンドやダイアログ API の要件セットは、さまざまな Office アプリケーションで共有されている API セットの例です。</span><span class="sxs-lookup"><span data-stu-id="abaf1-122">Add-in commands and Dialog API requirement sets are examples of API sets that the various Office applications share in common.</span></span>

<span data-ttu-id="abaf1-p106">アドインは、そのアドインが動作している Office アプリケーションのバージョンでサポートされている要件セットにある API のみを使用できます。 特定の Office アプリケーションのバージョンで使用できる要件セットを正確に確認するには、アプリケーション固有の要件セットに関する次の記事を参照してください。</span><span class="sxs-lookup"><span data-stu-id="abaf1-p106">An add-in can only use APIs in requirement sets that are supported by the version of Office application where the add-in is running. To know exactly which requirement sets are available for a specific Office application version, refer to the following application-specific requirement set articles:</span></span>

- <span data-ttu-id="abaf1-125">[Excel JavaScript API 要件セット](../reference/requirement-sets/excel-api-requirement-sets.md) (ExcelApi)</span><span class="sxs-lookup"><span data-stu-id="abaf1-125">[Excel JavaScript API requirement sets](../reference/requirement-sets/excel-api-requirement-sets.md) (ExcelApi)</span></span>
- <span data-ttu-id="abaf1-126">[Word JavaScript API 要件セット](../reference/requirement-sets/word-api-requirement-sets.md) (WordApi)</span><span class="sxs-lookup"><span data-stu-id="abaf1-126">[Word JavaScript API requirement sets](../reference/requirement-sets/word-api-requirement-sets.md) (WordApi)</span></span>
- <span data-ttu-id="abaf1-127">[OneNote JavaScript API 要件セット](../reference/requirement-sets/onenote-api-requirement-sets.md) (OneNoteApi)</span><span class="sxs-lookup"><span data-stu-id="abaf1-127">[OneNote JavaScript API requirement sets](../reference/requirement-sets/onenote-api-requirement-sets.md) (OneNoteApi)</span></span>
- <span data-ttu-id="abaf1-128">[PowerPoint JavaScript API 要件セット](../reference/requirement-sets/powerpoint-api-requirement-sets.md) (PowerPointApi)</span><span class="sxs-lookup"><span data-stu-id="abaf1-128">[PowerPoint JavaScript API requirement sets](../reference/requirement-sets/powerpoint-api-requirement-sets.md) (PowerPointApi)</span></span>
- <span data-ttu-id="abaf1-129">[Outlook API 要件セットについて](../reference/requirement-sets/outlook-api-requirement-sets.md) (Mailbox)</span><span class="sxs-lookup"><span data-stu-id="abaf1-129">[Understanding Outlook API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md) (Mailbox)</span></span>

<span data-ttu-id="abaf1-p107">一部の要件セットには、どの Office アプリケーションでも使用できる API が含まれています。 それらの要件セットの詳細については、次の記事を参照してください。</span><span class="sxs-lookup"><span data-stu-id="abaf1-p107">Some requirement sets contain APIs that can be used by any Office application. For information about these requirement sets, refer to the following articles:</span></span>

- [<span data-ttu-id="abaf1-132">Office の共通要件セット</span><span class="sxs-lookup"><span data-stu-id="abaf1-132">Office common requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="abaf1-133">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="abaf1-133">Add-in commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="abaf1-134">ダイアログ API の要件セット</span><span class="sxs-lookup"><span data-stu-id="abaf1-134">Dialog API requirement sets</span></span>](../reference/requirement-sets/dialog-api-requirement-sets.md)
- [<span data-ttu-id="abaf1-135">Identity API の要件セット</span><span class="sxs-lookup"><span data-stu-id="abaf1-135">Identity API requirement sets</span></span>](../reference/requirement-sets/identity-api-requirement-sets.md)

<span data-ttu-id="abaf1-p108">`ExcelApi 1.1` の "1.1" など、要件セットのバージョン番号は Office アプリケーションを基準にしています。 特定の要件セットのバージョン番号 (例: `ExcelApi 1.1`) は、Office.js のバージョン番号には対応しておらず、他の Office アプリケーション (Word、Outlook など) の要件セットにも対応していません。  Office アプリケーションの要件セットがリリースされる割合は、アプリケーションによって異なります。 たとえば、`ExcelApi 1.5` の方が `WordApi 1.3` 要件セットより前にリリースされました。</span><span class="sxs-lookup"><span data-stu-id="abaf1-p108">The version number of a requirement set, such as the "1.1" in `ExcelApi 1.1`, is relative to the Office application. The version number of a given requirement set (e.g., `ExcelApi 1.1`) does not correspond to the version number of Office.js or to requirement sets for other Office applications (e.g., Word, Outlook, etc.).  Requirement sets for the different Office applications are released at different rates. For example, `ExcelApi 1.5` was released before the `WordApi 1.3` requirement set.</span></span>


<span data-ttu-id="abaf1-140">Office JavaScript API ライブラリ (Office.js) には、現在利用可能なすべての要件セットが含まれています。</span><span class="sxs-lookup"><span data-stu-id="abaf1-140">The Office JavaScript API library (Office.js) includes all requirement sets that are currently available.</span></span> <span data-ttu-id="abaf1-141">`ExcelApi 1.3` や `WordApi 1.3` のような要件セットは存在しますが、`Office.js 1.3` のような要件セットは存在しません。</span><span class="sxs-lookup"><span data-stu-id="abaf1-141">While there is such a thing as requirement sets `ExcelApi 1.3` and `WordApi 1.3`, there is no `Office.js 1.3` requirement set.</span></span> <span data-ttu-id="abaf1-142">Office.js の最新リリースは、コンテンツ配信ネットワーク (CDN) 経由で配信される単一の Office エンドポイントとして維持されます。</span><span class="sxs-lookup"><span data-stu-id="abaf1-142">The latest release of Office.js is maintained as a single Office endpoint delivered via the content delivery network (CDN).</span></span> <span data-ttu-id="abaf1-143">バージョン管理や下位互換性の処理方法など、Office.js CDN に関する詳細については、「[Office JavaScript API について](../develop/understanding-the-javascript-api-for-office.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="abaf1-143">For more details around the Office.js CDN, including how versioning and backward compatibility is handled, see [Understanding the Office JavaScript API](../develop/understanding-the-javascript-api-for-office.md).</span></span>

## <a name="specify-office-applications-and-requirement-sets"></a><span data-ttu-id="abaf1-144">Office アプリケーションと要件セットを指定する</span><span class="sxs-lookup"><span data-stu-id="abaf1-144">Specify Office applications and requirement sets</span></span>

<span data-ttu-id="abaf1-p110">アドインに必要となる Office アプリケーションと要件セットは、さまざまな方法で指定できます。  詳細については、「[Office アプリケーションと API の要件を指定する](../develop/specify-office-hosts-and-api-requirements.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="abaf1-p110">There are various ways to specify which Office applications and requirement sets are required by an add-in.  For detailed information, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md)</span></span>

## <a name="see-also"></a><span data-ttu-id="abaf1-147">関連項目</span><span class="sxs-lookup"><span data-stu-id="abaf1-147">See also</span></span>

- [<span data-ttu-id="abaf1-148">Office アプリケーションと API の要件を指定する</span><span class="sxs-lookup"><span data-stu-id="abaf1-148">Specify Office applications and API requirements</span></span>](../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="abaf1-149">Office の最新バージョンをインストールする</span><span class="sxs-lookup"><span data-stu-id="abaf1-149">Install the latest version of Office</span></span>](../develop/install-latest-office-version.md)
- [<span data-ttu-id="abaf1-150">Microsoft 365 Apps 用更新プログラム チャネルの概要</span><span class="sxs-lookup"><span data-stu-id="abaf1-150">Overview of update channels for Microsoft 365 Apps</span></span>](/deployoffice/overview-of-update-channels-for-office-365-proplus)
- [<span data-ttu-id="abaf1-151">Microsoft 365 と Microsoft Teams による生産性の再構築</span><span class="sxs-lookup"><span data-stu-id="abaf1-151">Reimagine productivity with Microsoft 365 and Microsoft Teams</span></span>](https://products.office.com/compare-all-microsoft-office-products?tab=2)
