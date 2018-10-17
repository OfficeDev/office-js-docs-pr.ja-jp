---
title: Office のバージョンと要件セット
description: ''
ms.date: 03/29/2018
ms.openlocfilehash: ac3ae4fa3eeca9cfbd56b15168fc39d67139680d
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505994"
---
# <a name="office-versions-and-requirement-sets"></a><span data-ttu-id="d98d8-102">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="d98d8-102">Office versions and requirement sets</span></span>

<span data-ttu-id="d98d8-p101">Office にはプラットフォームやバージョンが異なるものが数多くあり、それらすべてが Office JavaScript API (Office.js) に含まれる API をすべてサポートしているわけではありません。このような状況に対処するため、Office アドインで必要な機能を Office ホストがサポートしているかどうかを判別するのに役立つ要件セットと呼ばれるシステムが用意されています。</span><span class="sxs-lookup"><span data-stu-id="d98d8-p101">There are many versions of Office on several platforms, and they don't all support every API in Office JavaScript API (Office.js). You may not always have control over the version of Office your users have installed.  To handle this situation, we provide a system called requirement sets to help you determine whether an Office host supports the capabilities you need in your Office Add-in.</span></span> 

> [!NOTE]
> - <span data-ttu-id="d98d8-106">Office は、Office for Windows、Office Online、Office for Mac、Office for iPad を含む複数のプラットフォームで実行できます。</span><span class="sxs-lookup"><span data-stu-id="d98d8-106">Office runs across multiple platforms, including Office for Windows, Office Online, Office for the Mac, and Office for the iPad.</span></span>  
> - <span data-ttu-id="d98d8-107">Office ホストの例は、Excel、Word、PowerPoint、Outlook、OneNote などの Office 製品です。</span><span class="sxs-lookup"><span data-stu-id="d98d8-107">Examples of Office hosts are Office Products: Excel, Word, PowerPoint, Outlook, OneNote, and so forth.</span></span>  
> - <span data-ttu-id="d98d8-108">要件セットとは、`ExcelApi 1.5` や `WordApi 1.3` などの、API メンバーの名前付きグループです。</span><span class="sxs-lookup"><span data-stu-id="d98d8-108">A requirement set is a named group of API members e.g., `ExcelApi 1.5`, `WordApi 1.3`, and so on.</span></span>  


## <a name="how-to-check-your-office-version"></a><span data-ttu-id="d98d8-109">Office のバージョンを確認する方法</span><span class="sxs-lookup"><span data-stu-id="d98d8-109">How to check your Office version</span></span>

<span data-ttu-id="d98d8-p102">使用している Office のバージョンを確認するには、Office アプリケーション内の **[ファイル]** メニューを選択し、**[アカウント]** を選択します。この Office のバージョンは、 **[製品情報]** セクションに表示されます。たとえば、次のスクリーンショットは Office バージョン 1802 (ビルド 9026.1000) を示しています</span><span class="sxs-lookup"><span data-stu-id="d98d8-p102">To identify the Office version that you're using, from within an Office application, select the **File** menu, and then choose **Account**. The version of Office will appear in the **Product Information** section. For example, the following screenshot indicates Office Version 1802 (Build 9026.1000):</span></span>

![Office のバージョン確認](../images/office-version-number-ui.jpg)


## <a name="office-requirement-sets-availability"></a><span data-ttu-id="d98d8-114">Office 要件セットの可用性</span><span class="sxs-lookup"><span data-stu-id="d98d8-114">Office requirement sets availability</span></span>

<span data-ttu-id="d98d8-p103">Office アドインは API 要件セットを使用して、使用する必要のある API メンバーを Office ホストがサポートしているかどうかを判別できます。要件セットのサポートは、Office ホストと Office ホストのバージョンによって異なります (前のセクションを参照してください)。</span><span class="sxs-lookup"><span data-stu-id="d98d8-p103">Office Add-ins can use API requirement sets to determine whether the Office host supports the API members that it need to use. Requirement set support varies by Office host and the Office host version (see previous section).</span></span>

<span data-ttu-id="d98d8-p104">一部の Office ホストでは、独自の API 要件セットがあります。たとえば、Excel API の最初の要件セットは `ExcelApi 1.1` で、Word API の最初の要件セットは `WordApi 1.1`でした。それ以降、追加の機能を提供するため、複数の新しい ExcelApi 要件セットと WordApi 要件セットが追加されています。</span><span class="sxs-lookup"><span data-stu-id="d98d8-p104">Some Office hosts have their own API requirement sets. For example, the first requirement set for the Excel API was `ExcelApi 1.1` and the first requirement set for the Word API was `WordApi 1.1`. Since then, multiple new ExcelApi requirement sets and WordApi requirement sets have been added to provide additional API functionality.</span></span>

<span data-ttu-id="d98d8-p105">さらに、アドイン コマンド (リボン機能拡張) やダイアログ ボックスを起動する機能 (ダイアログ API) など、他の機能が一般的な API に追加されました。アドイン コマンドやダイアログ API の要件セットは、さまざまな Office ホストで共有されている API セットの例です。</span><span class="sxs-lookup"><span data-stu-id="d98d8-p105">In addition, other functionality such as add-in commands (ribbon extensibility) and the ability to launch dialog boxes (Dialog API) were added to the common API. Add-in commands and Dialog API requirement sets are examples of API sets that the various Office hosts share in common.</span></span>

<span data-ttu-id="d98d8-p106">アドインは、そのアドインが動作している Office ホストのバージョンでサポートしている要件セットにある API のみを使用できます。特定の Office ホストのバージョンで使用できる要件セットを正確に確認するには、ホスト固有の要件セットに関する次の記事を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d98d8-p106">An add-in can only use APIs in requirement sets that are supported by the version of Office host where the add-in is running. To know exactly which requirement sets are available for a specific Office host version, refer to the following host-specific requirement set articles:</span></span>

- <span data-ttu-id="d98d8-124">[Excel JavaScript API 要件セット](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets?view=office-js) (ExcelApi)</span><span class="sxs-lookup"><span data-stu-id="d98d8-124">[Excel JavaScript API requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets?view=office-js) (ExcelApi)</span></span>
- <span data-ttu-id="d98d8-125">[Word JavaScript API 要件セット](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets?view=office-js) (WordApi)</span><span class="sxs-lookup"><span data-stu-id="d98d8-125">[Word JavaScript API requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets?view=office-js) (WordApi)</span></span>
- <span data-ttu-id="d98d8-126">[OneNote JavaScript API 要件セット](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets?view=office-js) (OneNoteApi)</span><span class="sxs-lookup"><span data-stu-id="d98d8-126">[OneNote JavaScript API requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets?view=office-js) (OneNoteApi)</span></span>
- <span data-ttu-id="d98d8-127">[Outlook API 要件セットについて](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets?view=office-js) (MailBox)</span><span class="sxs-lookup"><span data-stu-id="d98d8-127">[Understanding Outlook API requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets?view=office-js) (MailBox)</span></span>

<span data-ttu-id="d98d8-p107">一部の要件セットには、どの Office ホストでも使用できる API が含まれています。これらの要件のセットの詳細については、次の記事を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d98d8-p107">Some requirement sets contain APIs that can be used by any Office host. For information about these requirement sets, refer to the following articles:</span></span>

- [<span data-ttu-id="d98d8-130">Office の共通要件セット</span><span class="sxs-lookup"><span data-stu-id="d98d8-130">Office common requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets?view=office-js)
- [<span data-ttu-id="d98d8-131">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="d98d8-131">Add-in commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets?view=office-js)
- [<span data-ttu-id="d98d8-132">ダイアログ API の要件セット</span><span class="sxs-lookup"><span data-stu-id="d98d8-132">Dialog API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets?view=office-js)
- [<span data-ttu-id="d98d8-133">Identity API の要件セット</span><span class="sxs-lookup"><span data-stu-id="d98d8-133">Identity API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js)

<span data-ttu-id="d98d8-p108">`ExcelApi 1.1`の「1.1」など要件セットのバージョン番号は、Office ホストを基準としています。 特定の要件セットのバージョン番号 (たとえば、`ExcelApi 1.1`) は、Office.js や Office ホスト (たとえば Word、Outlook) の要件セットに対応しておらず、他の Office ホストの要件セットは、異なる時期にリリースされています。たとえば`ExcelApi 1.5` は`WordApi 1.3` 要件セットよりも前にリリースされました。</span><span class="sxs-lookup"><span data-stu-id="d98d8-p108">The version number of a requirement set, such as the "1.1" in `ExcelApi 1.1`, is relative to the Office host. The version number of a given requirement set (e.g., `ExcelApi 1.1`) does not correspond to the version number of Office.js or to requirement sets for other Office hosts (e.g., Word, Outlook, etc.).  Requirement sets for the different Office hosts are released at different speeds and times. For example, `ExcelApi 1.5` was released before the `WordApi 1.3` requirement set.</span></span>

<span data-ttu-id="d98d8-p109">JavaScript API for Office ライブラリ (Office.js) には現在利用できるすべての要件セットが含まれています。要件セット `ExcelApi 1.3` や `WordApi 1.3` がある一方で、 `Office.js 1.3` 要件セットはありません。最新リリースの Office.js は、コンテンツ配信ネットワーク (CDN) を介して配信される単一 Office エンドポイントとして維持されます。バージョン管理と下位互換性の処理方法など、Office.js CDN に関する詳細は、「 [JavaScript API for Office を理解する](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d98d8-p109">The JavaScript API for Office library (Office.js) includes all requirement sets that are currently available. While there is such a thing as requirement sets `ExcelApi 1.3` and `WordApi 1.3`, there is no `Office.js 1.3` requirement set. The latest release of Office.js is maintained as a single Office endpoint delivered via the content delivery network (CDN). For more details around the Office.js CDN, including how versioning and backward compatibility is handled, see [Understanding the JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span></span>

## <a name="specify-office-hosts-and-requirement-sets"></a><span data-ttu-id="d98d8-142">Office ホストと要件セットを指定する</span><span class="sxs-lookup"><span data-stu-id="d98d8-142">Specify Office hosts and requirement sets</span></span>

<span data-ttu-id="d98d8-p110">アドインに必要となる Office ホストと要件セットは、さまざまな方法で指定できます。詳細については、「 [Office のホストと API の要件を指定する](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d98d8-p110">There are various ways to specify which Office hosts and requirement sets are required by an add-in.  For detailed information, see [Specify Office hosts and API requirements](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)</span></span>


## <a name="see-also"></a><span data-ttu-id="d98d8-145">関連項目</span><span class="sxs-lookup"><span data-stu-id="d98d8-145">See also</span></span>

- [<span data-ttu-id="d98d8-146">Office のホストと API の要件を指定する</span><span class="sxs-lookup"><span data-stu-id="d98d8-146">Specify Office hosts and API requirements</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="d98d8-147">Office の最新バージョンをインストールする</span><span class="sxs-lookup"><span data-stu-id="d98d8-147">Install the latest version of Office</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/install-latest-office-version)
- [<span data-ttu-id="d98d8-148">Office 365 ProPlus 更新チャネルの概要</span><span class="sxs-lookup"><span data-stu-id="d98d8-148">Overview of update channels for Office 365 ProPlus</span></span>](https://docs.microsoft.com/deployoffice/overview-of-update-channels-for-office-365-proplus)
- [<span data-ttu-id="d98d8-149">Office 365 で Office を最大限に活用する</span><span class="sxs-lookup"><span data-stu-id="d98d8-149">Get the most from Office with Office 365</span></span>](https://products.office.com/compare-all-microsoft-office-products?tab=2)
