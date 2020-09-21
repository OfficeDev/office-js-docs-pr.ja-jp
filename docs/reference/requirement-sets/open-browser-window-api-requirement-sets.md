---
title: ブラウザーウィンドウの要件セットを開く
description: OpenBrowserWindow API をサポートする Office プラットフォームとビルドを指定します。
ms.date: 09/16/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 8f6966f5bdcecd9c55a20f2d640d066906c1b6a3
ms.sourcegitcommit: 2479812e677d1a7337765fe8f1c8345061d4091a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/19/2020
ms.locfileid: "48135239"
---
# <a name="open-browser-window-api-requirement-sets"></a><span data-ttu-id="43e6b-103">ブラウザーウィンドウ API の要件セットを開く</span><span class="sxs-lookup"><span data-stu-id="43e6b-103">Open Browser Window API requirement sets</span></span>

<span data-ttu-id="43e6b-p101">要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="43e6b-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="43e6b-107">OpenBrowserWindow API セットを使用すると、アドイン自体でサンドボックス内の webview コントロールでは実行できないタスクを実行するために、ブラウザーを開いておくことができます。たとえば、Microsoft Edge で webview コントロールが提供されている場合は、PDF ファイルをダウンロードします。</span><span class="sxs-lookup"><span data-stu-id="43e6b-107">The OpenBrowserWindow API set enables add-ins to open a browser to accomplish tasks that cannot always be done in the sandboxed webview control within the add-in itself; for example, downloading a PDF file when the webview control is provided by Microsoft Edge.</span></span>

<span data-ttu-id="43e6b-108">Office アドインは Office の複数のバージョンで機能します。</span><span class="sxs-lookup"><span data-stu-id="43e6b-108">Office Add-ins run across multiple versions of Office.</span></span> <span data-ttu-id="43e6b-109">次の表に、OpenBrowserWindow API の要件セット、その要件セットをサポートする Office ホストアプリケーション、Office アプリケーションのビルド番号またはバージョン番号を示します。</span><span class="sxs-lookup"><span data-stu-id="43e6b-109">The following table lists the OpenBrowserWindow API requirement sets, the Office host applications that support that requirement set, and the build or version numbers for the Office application.</span></span>

|  <span data-ttu-id="43e6b-110">要件セット</span><span class="sxs-lookup"><span data-stu-id="43e6b-110">Requirement set</span></span>  | <span data-ttu-id="43e6b-111">Windows 以降の Office 2013</span><span class="sxs-lookup"><span data-stu-id="43e6b-111">Office 2013 on Windows or later</span></span><br><span data-ttu-id="43e6b-112">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="43e6b-112">(one-time purchase)</span></span> | <span data-ttu-id="43e6b-113">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="43e6b-113">Office on Windows</span></span><br><span data-ttu-id="43e6b-114">(Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="43e6b-114">(connected to Office 365 subscription)</span></span> |  <span data-ttu-id="43e6b-115">Office on iPad</span><span class="sxs-lookup"><span data-stu-id="43e6b-115">Office on iPad</span></span><br><span data-ttu-id="43e6b-116">(Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="43e6b-116">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="43e6b-117">Office on Mac</span><span class="sxs-lookup"><span data-stu-id="43e6b-117">Office on Mac</span></span><br><span data-ttu-id="43e6b-118">(Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="43e6b-118">(connected to Office 365 subscription)</span></span>  | <span data-ttu-id="43e6b-119">Office on the web</span><span class="sxs-lookup"><span data-stu-id="43e6b-119">Office on the web</span></span>  |  <span data-ttu-id="43e6b-120">Office Online Server</span><span class="sxs-lookup"><span data-stu-id="43e6b-120">Office Online Server</span></span>  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="43e6b-121">OpenBrowserWindowApi 1.1</span><span class="sxs-lookup"><span data-stu-id="43e6b-121">OpenBrowserWindowApi 1.1</span></span>  | <span data-ttu-id="43e6b-122">N/A</span><span class="sxs-lookup"><span data-stu-id="43e6b-122">N/A</span></span> | <span data-ttu-id="43e6b-123">バージョン 1810 (ビルド 16.0.11001.20074) 以降</span><span class="sxs-lookup"><span data-stu-id="43e6b-123">Version 1810 (Build 16.0.11001.20074) or later</span></span> | <span data-ttu-id="43e6b-124">16.0.0.0 以降</span><span class="sxs-lookup"><span data-stu-id="43e6b-124">16.0.0.0 or later</span></span> | <span data-ttu-id="43e6b-125">16.0.0.0 以降</span><span class="sxs-lookup"><span data-stu-id="43e6b-125">16.0.0.0 or later</span></span> | <span data-ttu-id="43e6b-126">N/A</span><span class="sxs-lookup"><span data-stu-id="43e6b-126">N/A</span></span> | <span data-ttu-id="43e6b-127">N/A</span><span class="sxs-lookup"><span data-stu-id="43e6b-127">N/A</span></span>|

<span data-ttu-id="43e6b-128">バージョン、ビルド番号、Office Online Server の詳細については以下を参照してください。</span><span class="sxs-lookup"><span data-stu-id="43e6b-128">To find out more about versions, build numbers, and Office Online Server, see:</span></span>

- [<span data-ttu-id="43e6b-129">Office 365 クライアントの更新プログラム チャネル リリースのバージョン番号およびビルド番号</span><span class="sxs-lookup"><span data-stu-id="43e6b-129">Version and build numbers of update channel releases for Office 365 clients</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [<span data-ttu-id="43e6b-130">使用している Office のバージョンを確認する方法</span><span class="sxs-lookup"><span data-stu-id="43e6b-130">What version of Office am I using?</span></span>](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [<span data-ttu-id="43e6b-131">Office 365 クライアント アプリケーションのバージョン番号およびビルド番号を確認することができます。</span><span class="sxs-lookup"><span data-stu-id="43e6b-131">Where you can find the version and build number for an Office 365 client application</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [<span data-ttu-id="43e6b-132">Office Online Server 概要</span><span class="sxs-lookup"><span data-stu-id="43e6b-132">Office Online Server overview</span></span>](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="43e6b-133">Office 共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="43e6b-133">Office Common API requirement sets</span></span>

<span data-ttu-id="43e6b-134">共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="43e6b-134">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="openbrowserwindowapi-11"></a><span data-ttu-id="43e6b-135">OpenBrowserWindowApi 1.1</span><span class="sxs-lookup"><span data-stu-id="43e6b-135">OpenBrowserWindowApi 1.1</span></span>

<span data-ttu-id="43e6b-136">OpenBrowserWindowApi 1.1 は、API の最初のバージョンです。</span><span class="sxs-lookup"><span data-stu-id="43e6b-136">The OpenBrowserWindowApi 1.1 is the first version of the API.</span></span> <span data-ttu-id="43e6b-137">API の詳細については、「 [Office. ui](/javascript/api/office/office.context.ui) リファレンス」のトピックを参照してください。</span><span class="sxs-lookup"><span data-stu-id="43e6b-137">For details about the API, see the [Office.context.ui](/javascript/api/office/office.context.ui) reference topic.</span></span>

## <a name="see-also"></a><span data-ttu-id="43e6b-138">関連項目</span><span class="sxs-lookup"><span data-stu-id="43e6b-138">See also</span></span>

- [<span data-ttu-id="43e6b-139">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="43e6b-139">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="43e6b-140">Office のホストと API の要件を指定する</span><span class="sxs-lookup"><span data-stu-id="43e6b-140">Specify Office hosts and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="43e6b-141">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="43e6b-141">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
