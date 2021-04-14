---
title: ブラウザー ウィンドウの要件セットを開く
description: openBrowserWindow API Officeサポートするプラットフォームとビルドを指定します。
ms.date: 04/09/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: dd15136b350d42ec49187e436142aaecbfe70f40
ms.sourcegitcommit: 841bcad3c6c5139fd0953707c0be73ce890fa463
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/13/2021
ms.locfileid: "51687434"
---
# <a name="open-browser-window-api-requirement-sets"></a><span data-ttu-id="974c7-103">ブラウザー ウィンドウ API の要件セットを開く</span><span class="sxs-lookup"><span data-stu-id="974c7-103">Open Browser Window API requirement sets</span></span>

<span data-ttu-id="974c7-p101">要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="974c7-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="974c7-107">OpenBrowserWindow API セットを使用すると、アドインはブラウザーを開き、アドイン自体のサンドボックス Web ビュー コントロールで常に実行できないタスクを実行できます。たとえば、Webview コントロールが Microsoft Edge によって提供されている場合に PDF ファイルをダウンロードします。</span><span class="sxs-lookup"><span data-stu-id="974c7-107">The OpenBrowserWindow API set enables add-ins to open a browser to accomplish tasks that cannot always be done in the sandboxed webview control within the add-in itself; for example, downloading a PDF file when the webview control is provided by Microsoft Edge.</span></span>

<span data-ttu-id="974c7-108">Office アドインは Office の複数のバージョンで機能します。</span><span class="sxs-lookup"><span data-stu-id="974c7-108">Office Add-ins run across multiple versions of Office.</span></span> <span data-ttu-id="974c7-109">次の表に、OpenBrowserWindow API 要件セット、その要件セットをサポートする Office ホスト アプリケーション、および Office アプリケーションのビルドまたはバージョン番号を示します。</span><span class="sxs-lookup"><span data-stu-id="974c7-109">The following table lists the OpenBrowserWindow API requirement sets, the Office host applications that support that requirement set, and the build or version numbers for the Office application.</span></span>

|  <span data-ttu-id="974c7-110">要件セット</span><span class="sxs-lookup"><span data-stu-id="974c7-110">Requirement set</span></span>  | <span data-ttu-id="974c7-111">Office 2013以降の Windows でのインストール</span><span class="sxs-lookup"><span data-stu-id="974c7-111">Office 2013 on Windows or later</span></span><br><span data-ttu-id="974c7-112">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="974c7-112">(one-time purchase)</span></span> | <span data-ttu-id="974c7-113">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="974c7-113">Office on Windows</span></span><br><span data-ttu-id="974c7-114">(Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="974c7-114">(connected to Microsoft 365 subscription)</span></span> |  <span data-ttu-id="974c7-115">Office on iPad</span><span class="sxs-lookup"><span data-stu-id="974c7-115">Office on iPad</span></span><br><span data-ttu-id="974c7-116">(Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="974c7-116">(connected to Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="974c7-117">Office on Mac</span><span class="sxs-lookup"><span data-stu-id="974c7-117">Office on Mac</span></span><br><span data-ttu-id="974c7-118">(Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="974c7-118">(connected to Microsoft 365 subscription)</span></span>  | <span data-ttu-id="974c7-119">Office on the web</span><span class="sxs-lookup"><span data-stu-id="974c7-119">Office on the web</span></span>  |  <span data-ttu-id="974c7-120">Office Online Server</span><span class="sxs-lookup"><span data-stu-id="974c7-120">Office Online Server</span></span>  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="974c7-121">OpenBrowserWindowApi 1.1</span><span class="sxs-lookup"><span data-stu-id="974c7-121">OpenBrowserWindowApi 1.1</span></span>  | <span data-ttu-id="974c7-122">該当なし</span><span class="sxs-lookup"><span data-stu-id="974c7-122">N/A</span></span> | <span data-ttu-id="974c7-123">バージョン 1810 (ビルド 16.0.11001.20074) 以降</span><span class="sxs-lookup"><span data-stu-id="974c7-123">Version 1810 (Build 16.0.11001.20074) or later</span></span> | <span data-ttu-id="974c7-124">16.0.0.0 以降</span><span class="sxs-lookup"><span data-stu-id="974c7-124">16.0.0.0 or later</span></span> | <span data-ttu-id="974c7-125">16.0.0.0 以降</span><span class="sxs-lookup"><span data-stu-id="974c7-125">16.0.0.0 or later</span></span> | <span data-ttu-id="974c7-126">該当なし</span><span class="sxs-lookup"><span data-stu-id="974c7-126">N/A</span></span> | <span data-ttu-id="974c7-127">該当なし</span><span class="sxs-lookup"><span data-stu-id="974c7-127">N/A</span></span>|

> [!NOTE]
> <span data-ttu-id="974c7-128">OpenBrowserWindowApi 要件セットは、次のようにのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="974c7-128">The OpenBrowserWindowApi requirement set is only available as follows:</span></span>
>
> - <span data-ttu-id="974c7-129">Excel、PowerPoint、Word: Windows、Mac、iPad</span><span class="sxs-lookup"><span data-stu-id="974c7-129">Excel, PowerPoint, Word: Windows, Mac, iPad</span></span>
> - <span data-ttu-id="974c7-130">Outlook: Windows、Mac</span><span class="sxs-lookup"><span data-stu-id="974c7-130">Outlook: Windows, Mac</span></span>

<span data-ttu-id="974c7-131">バージョン、ビルド番号、Office Online Server の詳細については以下を参照してください。</span><span class="sxs-lookup"><span data-stu-id="974c7-131">To find out more about versions, build numbers, and Office Online Server, see:</span></span>

- [<span data-ttu-id="974c7-132">Microsoft 365 Apps の更新プログラム チャネル リリースのバージョン番号とビルド番号</span><span class="sxs-lookup"><span data-stu-id="974c7-132">Version and build numbers of update channel releases for Microsoft 365 Apps</span></span>](/officeupdates/update-history-microsoft365-apps-by-date)
- [<span data-ttu-id="974c7-133">使用している Office のバージョンを確認する方法</span><span class="sxs-lookup"><span data-stu-id="974c7-133">What version of Office am I using?</span></span>](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [<span data-ttu-id="974c7-134">クライアント アプリケーションのバージョンとビルド番号をOffice場所</span><span class="sxs-lookup"><span data-stu-id="974c7-134">Where you can find the version and build number for an Office client application</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [<span data-ttu-id="974c7-135">Office Online Server 概要</span><span class="sxs-lookup"><span data-stu-id="974c7-135">Office Online Server overview</span></span>](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="974c7-136">Office 共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="974c7-136">Office Common API requirement sets</span></span>

<span data-ttu-id="974c7-137">共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="974c7-137">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="openbrowserwindowapi-11"></a><span data-ttu-id="974c7-138">OpenBrowserWindowApi 1.1</span><span class="sxs-lookup"><span data-stu-id="974c7-138">OpenBrowserWindowApi 1.1</span></span>

<span data-ttu-id="974c7-139">OpenBrowserWindowApi 1.1 は API の最初のバージョンです。</span><span class="sxs-lookup"><span data-stu-id="974c7-139">The OpenBrowserWindowApi 1.1 is the first version of the API.</span></span> <span data-ttu-id="974c7-140">API の詳細については [、「Office.context.ui」を](/javascript/api/office/office.context#ui) 参照してください。</span><span class="sxs-lookup"><span data-stu-id="974c7-140">For details about the API, see the [Office.context.ui](/javascript/api/office/office.context#ui) reference topic.</span></span>

## <a name="see-also"></a><span data-ttu-id="974c7-141">関連項目</span><span class="sxs-lookup"><span data-stu-id="974c7-141">See also</span></span>

- [<span data-ttu-id="974c7-142">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="974c7-142">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="974c7-143">Office のホストと API の要件を指定する</span><span class="sxs-lookup"><span data-stu-id="974c7-143">Specify Office hosts and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="974c7-144">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="974c7-144">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
