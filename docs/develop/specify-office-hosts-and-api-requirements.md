---
title: Office のホストと API の要件を指定する
description: アドインが期待Office動作するアプリケーションと API 要件を指定する方法について説明します。
ms.date: 05/04/2021
localization_priority: Normal
ms.openlocfilehash: 07f2505dcfb16bf7000dca01a6d600aac9a63fa0
ms.sourcegitcommit: 8fbc7c7eb47875bf022e402b13858695a8536ec5
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/06/2021
ms.locfileid: "52253355"
---
# <a name="specify-office-applications-and-api-requirements"></a><span data-ttu-id="77bb4-103">Office アプリケーションと API 要件を指定する</span><span class="sxs-lookup"><span data-stu-id="77bb4-103">Specify Office applications and API requirements</span></span>

<span data-ttu-id="77bb4-104">アドインOfficeは、特定の Office アプリケーション、要件セット、API メンバー、または API のバージョンに依存して、期待通り動作する場合があります。</span><span class="sxs-lookup"><span data-stu-id="77bb4-104">Your Office Add-in might depend on a specific Office application, a requirement set, an API member, or a version of the API in order to work as expected.</span></span> <span data-ttu-id="77bb4-105">たとえば、次のようなアドインがあります。</span><span class="sxs-lookup"><span data-stu-id="77bb4-105">For example, your add-in might:</span></span>

- <span data-ttu-id="77bb4-106">1 つの Office アプリケーション (例: Word または Excel)、またはいくつかのアプリケーションで実行します。</span><span class="sxs-lookup"><span data-stu-id="77bb4-106">Run in a single Office application (e.g., Word or Excel), or several applications.</span></span>

- <span data-ttu-id="77bb4-p102">Office の一部のバージョンでのみ利用できる JavaScript API を使用します。たとえば、Excel 2016 で実行するアドインでは、Excel JavaScript API を使用することがあります。</span><span class="sxs-lookup"><span data-stu-id="77bb4-p102">Make use of JavaScript APIs that are only available in some versions of Office. For example, you might use the Excel JavaScript APIs in an add-in that runs in Excel 2016.</span></span>

- <span data-ttu-id="77bb4-109">アドインが使用する API メンバーをサポートするバージョンの Office でのみ実行します。</span><span class="sxs-lookup"><span data-stu-id="77bb4-109">Run only in versions of Office that support API members that your add-in uses.</span></span>

<span data-ttu-id="77bb4-110">この記事は、アドインが期待どおりに機能し、できるだけ多くのユーザーが利用できるようにするために選択する必要のあるオプションについて理解するのに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="77bb4-110">This article helps you understand which options you should choose to ensure that your add-in works as expected and reaches the broadest audience possible.</span></span>

> [!NOTE]
> <span data-ttu-id="77bb4-111">Office アドインが現在サポートされている場所の詳細なビューについては[、「Office](../overview/office-add-in-availability.md)クライアント アプリケーションと Office アドインのプラットフォームの可用性」ページを参照してください。</span><span class="sxs-lookup"><span data-stu-id="77bb4-111">For a high-level view of where Office Add-ins are currently supported, see the [Office client application and platform availability for Office Add-ins](../overview/office-add-in-availability.md) page.</span></span>

<span data-ttu-id="77bb4-112">この記事で説明する中心的な概念を次の表に示します。</span><span class="sxs-lookup"><span data-stu-id="77bb4-112">The following table lists core concepts discussed throughout this article.</span></span>

|<span data-ttu-id="77bb4-113">**概念**</span><span class="sxs-lookup"><span data-stu-id="77bb4-113">**Concept**</span></span>|<span data-ttu-id="77bb4-114">**説明**</span><span class="sxs-lookup"><span data-stu-id="77bb4-114">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="77bb4-115">Office アプリケーション、Office クライアント アプリケーション</span><span class="sxs-lookup"><span data-stu-id="77bb4-115">Office application, Office client application</span></span>|<span data-ttu-id="77bb4-p103">アドインの実行に使用される Office アプリケーション。たとえば、Word や Excel など。</span><span class="sxs-lookup"><span data-stu-id="77bb4-p103">The Office application used to run your add-in. For example, Word, Excel, and so on.</span></span>|
|<span data-ttu-id="77bb4-118">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="77bb4-118">Platform</span></span>|<span data-ttu-id="77bb4-119">ブラウザーやOfficeなど、アプリケーションが実行される場所iPad。</span><span class="sxs-lookup"><span data-stu-id="77bb4-119">Where the Office application runs, such as in a browser or on an iPad.</span></span>|
|<span data-ttu-id="77bb4-120">要件セット</span><span class="sxs-lookup"><span data-stu-id="77bb4-120">Requirement set</span></span>|<span data-ttu-id="77bb4-121">関連する API メンバーの名前付きグループ。</span><span class="sxs-lookup"><span data-stu-id="77bb4-121">A named group of related API members.</span></span> <span data-ttu-id="77bb4-122">アドインは要件セットを使用して、Officeで使用される API メンバーをサポートするかどうかを判断します。</span><span class="sxs-lookup"><span data-stu-id="77bb4-122">Add-ins use requirement sets to determine whether the Office application supports API members used by your add-in.</span></span> <span data-ttu-id="77bb4-123">個々の API メンバーのサポートをテストするよりも、要件セットのサポートをテストするほうが簡単です。</span><span class="sxs-lookup"><span data-stu-id="77bb4-123">It's easier to test for the support of a requirement set than for the support of individual API members.</span></span> <span data-ttu-id="77bb4-124">要件セットのサポートは、アプリケーションOfficeアプリケーションのバージョンによってOfficeされます。</span><span class="sxs-lookup"><span data-stu-id="77bb4-124">Requirement set support varies by Office application and the version of the Office application.</span></span> <br ><span data-ttu-id="77bb4-125">要件セットはマニフェスト ファイルで指定されます。</span><span class="sxs-lookup"><span data-stu-id="77bb4-125">Requirement sets are specified in the manifest file.</span></span> <span data-ttu-id="77bb4-126">マニフェストで要件セットを指定する場合は、アドインを実行するために Office アプリケーションが提供する必要がある API サポートの最小レベルを設定します。</span><span class="sxs-lookup"><span data-stu-id="77bb4-126">When you specify requirement sets in the manifest, you set the minimum level of API support that the Office application must provide in order to run your add-in.</span></span> <span data-ttu-id="77bb4-127">Officeで指定された要件セットをサポートしないアプリケーションではアドインを実行できないので、アドインは [マイ アドイン] に<span class="ui">表示されません</span>。これにより、アドインを使用できる場所が制限されます。</span><span class="sxs-lookup"><span data-stu-id="77bb4-127">Office applications that don't support requirement sets specified in the manifest can't run your add-in, and your add-in won't display in <span class="ui">My Add-ins</span>. This restricts where your add-in is available.</span></span> <span data-ttu-id="77bb4-128">コードでは、ランタイム チェックを使用します。</span><span class="sxs-lookup"><span data-stu-id="77bb4-128">In code using runtime checks.</span></span> <span data-ttu-id="77bb4-129">要件セットの詳細な一覧については、「[Office アドインの要件セット](../reference/requirement-sets/office-add-in-requirement-sets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="77bb4-129">For the complete list of requirement sets, see [Office Add-in requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md).</span></span>|
|<span data-ttu-id="77bb4-130">ランタイム チェック</span><span class="sxs-lookup"><span data-stu-id="77bb4-130">Runtime check</span></span>|<span data-ttu-id="77bb4-131">アドインを実行している Officeがアドインで使用される要件セットまたはメソッドをサポートするかどうかを判断するために実行時に実行されるテスト。</span><span class="sxs-lookup"><span data-stu-id="77bb4-131">A test that is performed at runtime to determine whether the Office application running your add-in supports requirement sets or methods used by your add-in.</span></span> <span data-ttu-id="77bb4-132">ランタイム チェックを実行するには、メソッド、要件セット、または要件セットの一部ではないメソッド名を持つ **if** ステートメント `isSetSupported` を使用します。</span><span class="sxs-lookup"><span data-stu-id="77bb4-132">To perform a runtime check, you use an **if** statement with the `isSetSupported` method, the requirement sets, or the method names that aren't part of a requirement set.</span></span> <span data-ttu-id="77bb4-133">ランタイム チェックを使用すると、アドインを、最も多くのお客様が利用できるものにできます。</span><span class="sxs-lookup"><span data-stu-id="77bb4-133">Use runtime checks to ensure that your add-in reaches the broadest number of customers.</span></span> <span data-ttu-id="77bb4-134">要件セットとは異なり、ランタイム チェックでは、Office アプリケーションがアドインを実行するために提供する必要がある最小レベルの API サポートは指定されません。</span><span class="sxs-lookup"><span data-stu-id="77bb4-134">Unlike requirement sets, runtime checks don't specify the minimum level of API support that the Office application must provide for your add-in to run.</span></span> <span data-ttu-id="77bb4-135">代わりに **、if** ステートメントを使用して、API メンバーがサポートされているかどうかを判断します。</span><span class="sxs-lookup"><span data-stu-id="77bb4-135">Instead, you use the **if** statement to determine whether an API member is supported.</span></span> <span data-ttu-id="77bb4-136">サポートされている場合には、アドインで追加機能を提供できます。</span><span class="sxs-lookup"><span data-stu-id="77bb4-136">If it is, you can provide additional functionality in your add-in.</span></span> <span data-ttu-id="77bb4-137">ランタイム チェックを使用するときは、自分のアドインは必ず **[個人用アドイン]** に表示されます。</span><span class="sxs-lookup"><span data-stu-id="77bb4-137">Your add-in will always display in **My Add-ins** when you use runtime checks.</span></span>|

## <a name="before-you-begin"></a><span data-ttu-id="77bb4-138">始める前に</span><span class="sxs-lookup"><span data-stu-id="77bb4-138">Before you begin</span></span>

<span data-ttu-id="77bb4-139">アドインで最新バージョンのアドイン マニフェスト スキーマを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="77bb4-139">Your add-in must use the most current version of the add-in manifest schema.</span></span> <span data-ttu-id="77bb4-140">アドインでランタイム チェックを使用する場合は、最新の JavaScript API (Office) ライブラリをoffice.jsしてください。</span><span class="sxs-lookup"><span data-stu-id="77bb4-140">If you use runtime checks in your add-in, ensure that you use the latest Office JavaScript API (office.js) library.</span></span>

### <a name="specify-the-latest-add-in-manifest-schema"></a><span data-ttu-id="77bb4-141">最新のアドイン マニフェスト スキーマを指定する</span><span class="sxs-lookup"><span data-stu-id="77bb4-141">Specify the latest add-in manifest schema</span></span>

<span data-ttu-id="77bb4-142">アドインのマニフェストでは、アドイン マニフェスト スキーマのバージョン 1.1 を使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="77bb4-142">Your add-in's manifest must use version 1.1 of the add-in manifest schema.</span></span> <span data-ttu-id="77bb4-143">アドイン マニフェスト [の OfficeApp](../reference/manifest/officeapp.md) 要素を次のように設定します。</span><span class="sxs-lookup"><span data-stu-id="77bb4-143">Set the [OfficeApp](../reference/manifest/officeapp.md) element in your add-in manifest as follows.</span></span> <span data-ttu-id="77bb4-144">次の使用例は、型を示 `TaskPaneApp` しています。</span><span class="sxs-lookup"><span data-stu-id="77bb4-144">This example shows the `TaskPaneApp` type.</span></span>

```XML
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
```

### <a name="specify-the-latest-office-javascript-api-library"></a><span data-ttu-id="77bb4-145">JavaScript API ライブラリOfficeを指定する</span><span class="sxs-lookup"><span data-stu-id="77bb4-145">Specify the latest Office JavaScript API library</span></span>

<span data-ttu-id="77bb4-146">ランタイム チェックを使用する場合は、コンテンツ配信ネットワーク (Office) から JavaScript API ライブラリの最新バージョンを参照CDN。</span><span class="sxs-lookup"><span data-stu-id="77bb4-146">If you use runtime checks, reference the most current version of the Office JavaScript API library from the content delivery network (CDN).</span></span> <span data-ttu-id="77bb4-147">その場合、HTML に次の `script` タグを追加します。</span><span class="sxs-lookup"><span data-stu-id="77bb4-147">To do this, add the following  `script` tag to your HTML.</span></span> <span data-ttu-id="77bb4-148">CDN URL で `/1/` を使用することで、Office.js の最新版が参照されます。</span><span class="sxs-lookup"><span data-stu-id="77bb4-148">Using `/1/` in the CDN URL ensures that you reference the most recent version of Office.js.</span></span>

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

## <a name="options-to-specify-office-applications-or-api-requirements"></a><span data-ttu-id="77bb4-149">アプリケーションまたは API Officeを指定するオプション</span><span class="sxs-lookup"><span data-stu-id="77bb4-149">Options to specify Office applications or API requirements</span></span>

<span data-ttu-id="77bb4-150">アプリケーションまたは API Officeを指定する場合、考慮すべきいくつかの要因があります。</span><span class="sxs-lookup"><span data-stu-id="77bb4-150">When you specify Office applications or API requirements, there are several factors to consider.</span></span> <span data-ttu-id="77bb4-151">次の図に、アドインで使用すべき手法の判別方法を示します。</span><span class="sxs-lookup"><span data-stu-id="77bb4-151">The following diagram shows how to decide which technique to use in your add-in.</span></span>

![アプリケーションまたは API の要件を指定するときに、アドインに最適なOfficeを選択する](../images/options-for-office-hosts.png)

- <span data-ttu-id="77bb4-153">アドインが 1 つのアプリケーションでOffice場合は、マニフェスト `Hosts` で要素を設定します。</span><span class="sxs-lookup"><span data-stu-id="77bb4-153">If your add-in runs in one Office application, set the `Hosts` element in the manifest.</span></span> <span data-ttu-id="77bb4-154">詳しくは、「[Hosts 要素を設定する](#set-the-hosts-element)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="77bb4-154">For more information, see [Set the Hosts element](#set-the-hosts-element).</span></span>

- <span data-ttu-id="77bb4-155">アドインを実行するためにOfficeアプリケーションでサポートする必要がある最小要件セットまたは API メンバーを設定するには、マニフェストで要素 `Requirements` を設定します。</span><span class="sxs-lookup"><span data-stu-id="77bb4-155">To set the minimum requirement set or API members that an Office application must support to run your add-in, set the `Requirements` element in the manifest.</span></span> <span data-ttu-id="77bb4-156">詳しくは、「[マニフェストで Requirements 要素を設定する](#set-the-requirements-element-in-the-manifest)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="77bb4-156">For more information, see [Set the Requirements element in the manifest](#set-the-requirements-element-in-the-manifest).</span></span>

- <span data-ttu-id="77bb4-157">特定の要件セットまたは API メンバーが Office アプリケーションで使用できる場合は、追加の機能を提供する場合は、アドインの JavaScript コードでランタイム チェックを実行します。</span><span class="sxs-lookup"><span data-stu-id="77bb4-157">If you would like to provide additional functionality if specific requirement sets or API members are available in the Office application, perform a runtime check in your add-in's JavaScript code.</span></span> <span data-ttu-id="77bb4-158">たとえば、アドインが Excel 2016 で機能する場合は、Excel JavaScript API の API メンバーを使用して追加の機能を提供します。</span><span class="sxs-lookup"><span data-stu-id="77bb4-158">For example, if your add-in runs in Excel 2016, use API members from the Excel JavaScript API to provide additional functionality.</span></span> <span data-ttu-id="77bb4-159">詳細については、「[JavaScript コードでランタイム チェックを使用する](#use-runtime-checks-in-your-javascript-code)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="77bb4-159">For more information, see [Use runtime checks in your JavaScript code](#use-runtime-checks-in-your-javascript-code).</span></span>

## <a name="set-the-hosts-element"></a><span data-ttu-id="77bb4-160">Hosts 要素を設定する</span><span class="sxs-lookup"><span data-stu-id="77bb4-160">Set the Hosts element</span></span>

<span data-ttu-id="77bb4-161">アドインを 1 つのクライアント アプリケーションでOfficeするには、マニフェストの `Hosts` and `Host` 要素を使用します。</span><span class="sxs-lookup"><span data-stu-id="77bb4-161">To make your add-in run in one Office client application, use the `Hosts` and `Host` elements in the manifest.</span></span> <span data-ttu-id="77bb4-162">要素を指定しない場合、アドインは、指定した種類 (メール、作業ウィンドウ、またはコンテンツ) でサポートされているすべての Office アプリケーションで `Hosts` `OfficeApp` 実行されます。</span><span class="sxs-lookup"><span data-stu-id="77bb4-162">If you don't specify the `Hosts` element, your add-in will run in all Office applications supported by the specified `OfficeApp` type (that is, Mail, Task pane, or Content).</span></span>

<span data-ttu-id="77bb4-163">たとえば、次の宣言と宣言は、アドインが Excel のすべてのリリース (Excel on the web、Windows、および iPad を含む) で動作 `Hosts` `Host` iPad。</span><span class="sxs-lookup"><span data-stu-id="77bb4-163">For example, the following `Hosts` and `Host` declaration specifies that the add-in will work with any release of Excel, which includes Excel on the web, Windows, and iPad.</span></span>

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

<span data-ttu-id="77bb4-164">要素 `Hosts` には、1 つ以上の要素を含 `Host` めできます。</span><span class="sxs-lookup"><span data-stu-id="77bb4-164">The `Hosts` element can contain one or more `Host` elements.</span></span> <span data-ttu-id="77bb4-165">要素 `Host` は、アドインOffice必要なアプリケーションを指定します。</span><span class="sxs-lookup"><span data-stu-id="77bb4-165">The `Host` element specifies the Office application your add-in requires.</span></span> <span data-ttu-id="77bb4-166">属性 `Name` は必須であり、次のいずれかの値に設定できます。</span><span class="sxs-lookup"><span data-stu-id="77bb4-166">The `Name` attribute is required and can be set to one of the following values.</span></span>

| <span data-ttu-id="77bb4-167">名前</span><span class="sxs-lookup"><span data-stu-id="77bb4-167">Name</span></span>          | <span data-ttu-id="77bb4-168">Office クライアント アプリケーション</span><span class="sxs-lookup"><span data-stu-id="77bb4-168">Office client applications</span></span>                     | <span data-ttu-id="77bb4-169">使用可能なアドインの種類</span><span class="sxs-lookup"><span data-stu-id="77bb4-169">Available add-in types</span></span> |
|:--------------|:-----------------------------------------------|:-----------------------|
| <span data-ttu-id="77bb4-170">データベース</span><span class="sxs-lookup"><span data-stu-id="77bb4-170">Database</span></span>      | <span data-ttu-id="77bb4-171">Access Web アプリ</span><span class="sxs-lookup"><span data-stu-id="77bb4-171">Access web apps</span></span>                                | <span data-ttu-id="77bb4-172">作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="77bb4-172">Task pane</span></span>              |
| <span data-ttu-id="77bb4-173">Document</span><span class="sxs-lookup"><span data-stu-id="77bb4-173">Document</span></span>      | <span data-ttu-id="77bb4-174">Word on the web、Windows、Mac、iPad</span><span class="sxs-lookup"><span data-stu-id="77bb4-174">Word on the web, Windows, Mac, iPad</span></span>            | <span data-ttu-id="77bb4-175">作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="77bb4-175">Task pane</span></span>              |
| <span data-ttu-id="77bb4-176">Mailbox</span><span class="sxs-lookup"><span data-stu-id="77bb4-176">Mailbox</span></span>       | <span data-ttu-id="77bb4-177">Outlook、Windows、Mac、Android、iOS</span><span class="sxs-lookup"><span data-stu-id="77bb4-177">Outlook on the web, Windows, Mac, Android, iOS</span></span> | <span data-ttu-id="77bb4-178">メール</span><span class="sxs-lookup"><span data-stu-id="77bb4-178">Mail</span></span>                   |
| <span data-ttu-id="77bb4-179">Notebook</span><span class="sxs-lookup"><span data-stu-id="77bb4-179">Notebook</span></span>      | <span data-ttu-id="77bb4-180">OneNote on the web</span><span class="sxs-lookup"><span data-stu-id="77bb4-180">OneNote on the web</span></span>                             | <span data-ttu-id="77bb4-181">作業ウィンドウ、コンテンツ</span><span class="sxs-lookup"><span data-stu-id="77bb4-181">Task pane, Content</span></span>     |
| <span data-ttu-id="77bb4-182">Presentation</span><span class="sxs-lookup"><span data-stu-id="77bb4-182">Presentation</span></span>  | <span data-ttu-id="77bb4-183">PowerPoint on the web、Windows、Mac、iPad</span><span class="sxs-lookup"><span data-stu-id="77bb4-183">PowerPoint on the web, Windows, Mac, iPad</span></span>      | <span data-ttu-id="77bb4-184">作業ウィンドウ、コンテンツ</span><span class="sxs-lookup"><span data-stu-id="77bb4-184">Task pane, Content</span></span>     |
| <span data-ttu-id="77bb4-185">Project</span><span class="sxs-lookup"><span data-stu-id="77bb4-185">Project</span></span>       | <span data-ttu-id="77bb4-186">Windows での Project</span><span class="sxs-lookup"><span data-stu-id="77bb4-186">Project on Windows</span></span>                             | <span data-ttu-id="77bb4-187">作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="77bb4-187">Task pane</span></span>              |
| <span data-ttu-id="77bb4-188">Workbook</span><span class="sxs-lookup"><span data-stu-id="77bb4-188">Workbook</span></span>      | <span data-ttu-id="77bb4-189">Excel on the web、Windows、Mac、iPad</span><span class="sxs-lookup"><span data-stu-id="77bb4-189">Excel on the web, Windows, Mac, iPad</span></span>           | <span data-ttu-id="77bb4-190">作業ウィンドウ、コンテンツ</span><span class="sxs-lookup"><span data-stu-id="77bb4-190">Task pane, Content</span></span>     |

> [!NOTE]
> <span data-ttu-id="77bb4-191">この `Name` 属性は、アドインOffice実行できるクライアント アプリケーションの名前を指定します。</span><span class="sxs-lookup"><span data-stu-id="77bb4-191">The `Name` attribute specifies the Office client application that can run your add-in.</span></span> <span data-ttu-id="77bb4-192">Officeアプリケーションは、さまざまなプラットフォームでサポートされ、デスクトップ、Web ブラウザー、タブレット、およびモバイル デバイスで実行されます。</span><span class="sxs-lookup"><span data-stu-id="77bb4-192">Office applications are supported on different platforms and run on desktops, web browsers, tablets, and mobile devices.</span></span> <span data-ttu-id="77bb4-193">アドインを実行するために使用するプラットフォームを指定することはできません。</span><span class="sxs-lookup"><span data-stu-id="77bb4-193">You can't specify which platform can be used to run your add-in.</span></span> <span data-ttu-id="77bb4-194">たとえば、指定した場合は、web OutlookとWindowsの両方をアドインの実行 `Mailbox` に使用できます。</span><span class="sxs-lookup"><span data-stu-id="77bb4-194">For example, if you specify `Mailbox`, both Outlook on the web and on Windows can be used to run your add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="77bb4-195">SharePoint で Access Web アプリとデータベースを作成して使用することは推奨されなくなりました。</span><span class="sxs-lookup"><span data-stu-id="77bb4-195">We no longer recommend that you create and use Access web apps and databases in SharePoint.</span></span> <span data-ttu-id="77bb4-196">代わりに、[Microsoft PowerApps](https://powerapps.microsoft.com/) を使用して、コード作成が不要な Web とモバイル デバイス用ビジネス ソリューションをビルドすることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="77bb4-196">As an alternative, we recommend that you use [Microsoft PowerApps](https://powerapps.microsoft.com/) to build no-code business solutions for web and mobile devices.</span></span>

## <a name="set-the-requirements-element-in-the-manifest"></a><span data-ttu-id="77bb4-197">マニフェストで Requirements 要素を設定する</span><span class="sxs-lookup"><span data-stu-id="77bb4-197">Set the Requirements element in the manifest</span></span>

<span data-ttu-id="77bb4-198">要素は、アドインを実行するために、Officeアプリケーションでサポートする必要がある最小要件セットまたは API メンバー `Requirements` を指定します。</span><span class="sxs-lookup"><span data-stu-id="77bb4-198">The `Requirements` element specifies the minimum requirement sets or API members that must be supported by the Office application to run your add-in.</span></span> <span data-ttu-id="77bb4-199">要素 `Requirements` は、アドインで使用される要件セットと個々のメソッドの両方を指定できます。</span><span class="sxs-lookup"><span data-stu-id="77bb4-199">The `Requirements` element can specify both requirement sets and individual methods used in your add-in.</span></span> <span data-ttu-id="77bb4-200">アドイン マニフェスト スキーマのバージョン 1.1 では、アドインを除くすべてのアドインの要素 `Requirements` Outlookです。</span><span class="sxs-lookup"><span data-stu-id="77bb4-200">In version 1.1 of the add-in manifest schema, the `Requirements` element is optional for all add-ins, except for Outlook add-ins.</span></span>

> [!WARNING]
> <span data-ttu-id="77bb4-201">要素を使用 `Requirements` して、アドインで使用する必要がある重要な要件セットまたは API メンバーのみを指定します。</span><span class="sxs-lookup"><span data-stu-id="77bb4-201">Only use the `Requirements` element to specify critical requirement sets or API members that your add-in must use.</span></span> <span data-ttu-id="77bb4-202">Office アプリケーションまたはプラットフォームが要素で指定された要件セットまたは API メンバーをサポートしない場合、アドインは、そのアプリケーションまたはプラットフォームでは実行されません。また、My アドインには `Requirements` **表示** されません。代わりに、Office アプリケーションのすべてのプラットフォーム (Excel on the web、Windows、iPad など) でアドインを使用iPad。</span><span class="sxs-lookup"><span data-stu-id="77bb4-202">If the Office application or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that application or platform, and won't display in **My Add-ins**. Instead, we recommend that you make your add-in available on all platforms of an Office application, such as Excel on the web, Windows, and iPad.</span></span> <span data-ttu-id="77bb4-203">すべてのアプリケーションとプラットフォームでアドインをOfficeするには、要素の代わりにランタイム チェックを使用 `Requirements` します。</span><span class="sxs-lookup"><span data-stu-id="77bb4-203">To make your add-in available on  _all_ Office applications and platforms, use runtime checks instead of the `Requirements` element.</span></span>

<span data-ttu-id="77bb4-204">次のコード例は、次をサポートしているすべてのクライアント アプリケーションでOfficeアドインを示しています。</span><span class="sxs-lookup"><span data-stu-id="77bb4-204">The following code example shows an add-in that loads in all Office client applications that support the following:</span></span>

-  <span data-ttu-id="77bb4-205">`TableBindings` 要件セット 。最小バージョンは "1.1" です。</span><span class="sxs-lookup"><span data-stu-id="77bb4-205">`TableBindings` requirement set, which has a minimum version of "1.1".</span></span>

-  <span data-ttu-id="77bb4-206">`OOXML` 要件セット 。最小バージョンは "1.1" です。</span><span class="sxs-lookup"><span data-stu-id="77bb4-206">`OOXML` requirement set, which has a minimum version of "1.1".</span></span>

-  <span data-ttu-id="77bb4-207">`Document.getSelectedDataAsync` メソッド。</span><span class="sxs-lookup"><span data-stu-id="77bb4-207">`Document.getSelectedDataAsync` method.</span></span>

```XML
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="TableBindings" MinVersion="1.1"/>
      <Set Name="OOXML" MinVersion="1.1"/>
   </Sets>
   <Methods>
      <Method Name="Document.getSelectedDataAsync"/>
   </Methods>
</Requirements>
```

- <span data-ttu-id="77bb4-208">要素 `Requirements` には、子要素 `Sets` と子 `Methods` 要素が含まれます。</span><span class="sxs-lookup"><span data-stu-id="77bb4-208">The `Requirements` element contains the `Sets` and `Methods` child elements.</span></span>

- <span data-ttu-id="77bb4-209">要素 `Sets` には、1 つ以上の要素を含 `Set` めできます。</span><span class="sxs-lookup"><span data-stu-id="77bb4-209">The `Sets` element can contain one or more `Set` elements.</span></span> <span data-ttu-id="77bb4-210">`DefaultMinVersion` すべての子要素の `MinVersion` 既定値を指定 `Set` します。</span><span class="sxs-lookup"><span data-stu-id="77bb4-210">`DefaultMinVersion` specifies the default `MinVersion` value of all child `Set` elements.</span></span>

- <span data-ttu-id="77bb4-211">要素 `Set` は、アドインを実行するためにOfficeアプリケーションがサポートする必要がある要件セットを指定します。</span><span class="sxs-lookup"><span data-stu-id="77bb4-211">The `Set` element specifies requirement sets that the Office application must support to run the add-in.</span></span> <span data-ttu-id="77bb4-212">属性 `Name` は、要件セットの名前を指定します。</span><span class="sxs-lookup"><span data-stu-id="77bb4-212">The `Name` attribute specifies the name of the requirement set.</span></span> <span data-ttu-id="77bb4-213">要件 `MinVersion` セットの最小バージョンを指定します。</span><span class="sxs-lookup"><span data-stu-id="77bb4-213">The `MinVersion` specifies the minimum version of the requirement set.</span></span> <span data-ttu-id="77bb4-214">`MinVersion`overrides の値 API メンバーが属する要件セットと要件セットのバージョンの詳細については、「Officeアドイン要件セット」 `DefaultMinVersion` [を参照してください](../reference/requirement-sets/office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="77bb4-214">`MinVersion` overrides the value of `DefaultMinVersion` For more information about requirement sets and requirement set versions that your API members belong to, see [Office Add-in requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md).</span></span>

- <span data-ttu-id="77bb4-215">要素 `Methods` には、1 つ以上の要素を含 `Method` めできます。</span><span class="sxs-lookup"><span data-stu-id="77bb4-215">The `Methods` element can contain one or more `Method` elements.</span></span> <span data-ttu-id="77bb4-216">アドインで要素を `Methods` 使用Outlookすることはできません。</span><span class="sxs-lookup"><span data-stu-id="77bb4-216">You can't use the `Methods` element with Outlook add-ins.</span></span>

- <span data-ttu-id="77bb4-217">要素は、アドインが実行されるアプリケーションでサポートされる必要Office個別 `Method` のメソッドを指定します。</span><span class="sxs-lookup"><span data-stu-id="77bb4-217">The `Method` element specifies an individual method that must be supported in the Office application where your add-in runs.</span></span> <span data-ttu-id="77bb4-218">属性 `Name` は必須であり、親オブジェクトで修飾されたメソッドの名前を指定します。</span><span class="sxs-lookup"><span data-stu-id="77bb4-218">The `Name` attribute is required and specifies the name of the method qualified with its parent object.</span></span>

## <a name="use-runtime-checks-in-your-javascript-code"></a><span data-ttu-id="77bb4-219">JavaScript コードでランタイム チェックを使用する</span><span class="sxs-lookup"><span data-stu-id="77bb4-219">Use runtime checks in your JavaScript code</span></span>

<span data-ttu-id="77bb4-220">特定の要件セットがアプリケーションでサポートされている場合は、アドインに追加の機能を提供Officeがあります。</span><span class="sxs-lookup"><span data-stu-id="77bb4-220">You might want to provide additional functionality in your add-in if certain requirement sets are supported by the Office application.</span></span> <span data-ttu-id="77bb4-221">たとえば、アドインで Word 2016 を実行する場合、既存のアドインで Word JavaScript API を使用することがあります。</span><span class="sxs-lookup"><span data-stu-id="77bb4-221">For example, you might want to use the Word JavaScript APIs in your existing add-in if your add-in runs in Word 2016.</span></span> <span data-ttu-id="77bb4-222">その場合、要件セットの名前を指定し、[isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="77bb4-222">To do this, you use the [isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) method with the name of the requirement set.</span></span> <span data-ttu-id="77bb4-223">`isSetSupported`実行時に、アドインを実行Officeアプリケーションが要件セットをサポートするかどうかを判断します。</span><span class="sxs-lookup"><span data-stu-id="77bb4-223">`isSetSupported` determines, at runtime, whether the Office application running the add-in supports the requirement set.</span></span> <span data-ttu-id="77bb4-224">要件セットがサポートされている場合は、true を返し、その要件セットの API メンバーを使用する追加のコード `isSetSupported` を実行します。 </span><span class="sxs-lookup"><span data-stu-id="77bb4-224">If the requirement set is supported, `isSetSupported` returns **true** and runs the additional code that uses the API members from that requirement set.</span></span> <span data-ttu-id="77bb4-225">アプリケーションがOfficeが要件セットをサポートしない場合 `isSetSupported` **、false** を返し、追加のコードは実行されません。</span><span class="sxs-lookup"><span data-stu-id="77bb4-225">If the Office application doesn't support the requirement set, `isSetSupported` returns **false** and the additional code won't run.</span></span> <span data-ttu-id="77bb4-226">次のコードは `isSetSupported` と共に使用する構文を示しています。</span><span class="sxs-lookup"><span data-stu-id="77bb4-226">The following code shows the syntax to use with `isSetSupported`.</span></span>

```js
if (Office.context.requirements.isSetSupported(RequirementSetName, MinimumVersion))
{
   // Code that uses API members from RequirementSetName.
}

```

- <span data-ttu-id="77bb4-227">_RequirementSetName_ (必須) は、要件セットの名前を表す文字列です (例: "**ExcelApi**"、"**Mailbox**" など)。</span><span class="sxs-lookup"><span data-stu-id="77bb4-227">_RequirementSetName_ (required) is a string that represents the name of the requirement set (e.g., "**ExcelApi**", "**Mailbox**", etc.).</span></span> <span data-ttu-id="77bb4-228">利用できる要件セットの詳細については、「[Office アドインの要件セット](../reference/requirement-sets/office-add-in-requirement-sets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="77bb4-228">For more information about available requirement sets, see [Office Add-in requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md).</span></span>
- <span data-ttu-id="77bb4-229">_MinimumVersion_ (省略可能) は、ステートメント内のコードを実行するために Office アプリケーションがサポートする必要がある最小要件セット のバージョンを指定する文字列 `if` です (たとえば **、"1.9")。**</span><span class="sxs-lookup"><span data-stu-id="77bb4-229">_MinimumVersion_ (optional) is a string that specifies the minimum requirement set version that the Office application must support in order for the code within the `if` statement to run (e.g., "**1.9**").</span></span>

> [!WARNING]
> <span data-ttu-id="77bb4-230">メソッドを呼び `isSetSupported` 出す場合、パラメーターの値 (指定されている場合 `MinimumVersion` ) は文字列である必要があります。</span><span class="sxs-lookup"><span data-stu-id="77bb4-230">When calling the `isSetSupported` method, the value of the `MinimumVersion` parameter (if specified) should be a string.</span></span> <span data-ttu-id="77bb4-231">これは、JavaScript パーサーでは、1.1 や 1.10 のような数値の間の差異を区別できないが、"1.1" や "1.10" などの文字列値ではできるからです。</span><span class="sxs-lookup"><span data-stu-id="77bb4-231">This is because the JavaScript parser cannot differentiate between numeric values such as 1.1 and 1.10, where as it can for string values such as "1.1" and "1.10".</span></span>
> <span data-ttu-id="77bb4-232">`number` のオーバーロードは非推奨になります。</span><span class="sxs-lookup"><span data-stu-id="77bb4-232">The `number` overload is deprecated.</span></span>

<span data-ttu-id="77bb4-233">次 `isSetSupported` のように、 `RequirementSetName` アプリケーションに関連付Office使用します。</span><span class="sxs-lookup"><span data-stu-id="77bb4-233">Use `isSetSupported` with the `RequirementSetName` associated with the Office application as follows.</span></span>

|<span data-ttu-id="77bb4-234">Office アプリケーション</span><span class="sxs-lookup"><span data-stu-id="77bb4-234">Office application</span></span>|<span data-ttu-id="77bb4-235">RequirementSetName</span><span class="sxs-lookup"><span data-stu-id="77bb4-235">RequirementSetName</span></span>|
|---|---|
|<span data-ttu-id="77bb4-236">Excel</span><span class="sxs-lookup"><span data-stu-id="77bb4-236">Excel</span></span>|<span data-ttu-id="77bb4-237">ExcelApi</span><span class="sxs-lookup"><span data-stu-id="77bb4-237">ExcelApi</span></span>|
|<span data-ttu-id="77bb4-238">OneNote</span><span class="sxs-lookup"><span data-stu-id="77bb4-238">OneNote</span></span>|<span data-ttu-id="77bb4-239">OneNoteApi</span><span class="sxs-lookup"><span data-stu-id="77bb4-239">OneNoteApi</span></span>|
|<span data-ttu-id="77bb4-240">Outlook</span><span class="sxs-lookup"><span data-stu-id="77bb4-240">Outlook</span></span>|<span data-ttu-id="77bb4-241">Mailbox</span><span class="sxs-lookup"><span data-stu-id="77bb4-241">Mailbox</span></span>|
|<span data-ttu-id="77bb4-242">Word</span><span class="sxs-lookup"><span data-stu-id="77bb4-242">Word</span></span>|<span data-ttu-id="77bb4-243">WordApi</span><span class="sxs-lookup"><span data-stu-id="77bb4-243">WordApi</span></span>|

<span data-ttu-id="77bb4-244">これらの `isSetSupported` アプリケーションのメソッドと要件セットは、アプリケーションの最新のOffice.jsで使用CDN。</span><span class="sxs-lookup"><span data-stu-id="77bb4-244">The `isSetSupported` method and the requirement sets for these applications are available in the latest Office.js file on the CDN.</span></span> <span data-ttu-id="77bb4-245">アドインから例外をOffice.js場合CDN、未定義のため、アドインで `isSetSupported` 例外が生成される場合があります。</span><span class="sxs-lookup"><span data-stu-id="77bb4-245">If you don't use Office.js from the CDN, your add-in might generate exceptions because `isSetSupported` will be undefined.</span></span> <span data-ttu-id="77bb4-246">詳細については[、「JavaScript API ライブラリの最新のOffice指定する」を参照してください](#specify-the-latest-office-javascript-api-library)。</span><span class="sxs-lookup"><span data-stu-id="77bb4-246">For more information, see [Specify the latest Office JavaScript API library](#specify-the-latest-office-javascript-api-library).</span></span>

<span data-ttu-id="77bb4-247">次のコード例は、アドインが異なる要件セットまたは API メンバーをサポートする可能性Officeアプリケーションに対して異なる機能を提供する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="77bb4-247">The following code example shows how an add-in can provide different functionality for different Office applications that might support different requirement sets or API members.</span></span>

```js
if (Office.context.requirements.isSetSupported('WordApi', '1.1'))
{
    // Run code that provides additional functionality using the Word JavaScript API when the add-in runs in Word 2016 or later.
}
else if (Office.context.requirements.isSetSupported('CustomXmlParts'))
{
    // Run code that uses API members from the CustomXmlParts requirement set.
}
else
{
    // Run additional code when the Office application is not Word 2016 or later and does not support the CustomXmlParts requirement set.
}

```

## <a name="runtime-checks-using-methods-not-in-a-requirement-set"></a><span data-ttu-id="77bb4-248">要件セットにないメソッドを使用したランタイム チェック</span><span class="sxs-lookup"><span data-stu-id="77bb4-248">Runtime checks using methods not in a requirement set</span></span>

<span data-ttu-id="77bb4-249">API の一部のメンバーは、要件のセットに属していません。</span><span class="sxs-lookup"><span data-stu-id="77bb4-249">Some API members don't belong to requirement sets.</span></span> <span data-ttu-id="77bb4-250">これは[、Office JavaScript API](../reference/javascript-api-for-office.md)名前空間の一部である API メンバー (Outlook メールボックス API を除くすべての API) にのみ適用されますが `Office.` [、Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) (内の何でも) [](/javascript/api/outlook) `Word.` [、Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) (内の何でも)、または OneNote `Excel.` [JavaScript API](../reference/overview/onenote-add-ins-javascript-reference.md) ( `OneNote.` 何でも) 名前空間に属する API メンバーには適用されません。</span><span class="sxs-lookup"><span data-stu-id="77bb4-250">This only applies to API members that are part of the [Office JavaScript API](../reference/javascript-api-for-office.md) namespace (anything under `Office.` except [Outlook Mailbox APIs](/javascript/api/outlook)), but not API members that belong to the [Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) (anything in `Word.`), [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) (anything in `Excel.`), or [OneNote JavaScript API](../reference/overview/onenote-add-ins-javascript-reference.md) (anything in `OneNote.`) namespaces.</span></span> <span data-ttu-id="77bb4-251">アドインが要件セットの一部ではないメソッドに依存している場合は、ランタイム チェックを使用して、次のコード例に示すように、メソッドが Office アプリケーションでサポートされているかどうかを判断できます。</span><span class="sxs-lookup"><span data-stu-id="77bb4-251">When your add-in depends on a method that is not part of a requirement set, you can use the runtime check to determine whether the method is supported by the Office application, as shown in the following code example.</span></span> <span data-ttu-id="77bb4-252">要件セットに属さないメソッドの詳細な一覧については、「[Office アドインの要件セット](../reference/requirement-sets/office-add-in-requirement-sets.md#methods-that-arent-part-of-a-requirement-set)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="77bb4-252">For a complete list of methods that don't belong to a requirement set, see [Office Add-in requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md#methods-that-arent-part-of-a-requirement-set).</span></span>

> [!NOTE]
> <span data-ttu-id="77bb4-253">アドインのコードでのこの種のランタイム チェックは、限定的に使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="77bb4-253">We recommend that you limit the use of this type of runtime check in your add-in's code.</span></span>

<span data-ttu-id="77bb4-254">次のコード例は、アプリケーションがサポートOfficeチェックします `document.setSelectedDataAsync` 。</span><span class="sxs-lookup"><span data-stu-id="77bb4-254">The following code example checks whether the Office application supports `document.setSelectedDataAsync`.</span></span>

```js
if (Office.context.document.setSelectedDataAsync)
{
    // Run code that uses `document.setSelectedDataAsync`.
}
```


## <a name="see-also"></a><span data-ttu-id="77bb4-255">関連項目</span><span class="sxs-lookup"><span data-stu-id="77bb4-255">See also</span></span>

- [<span data-ttu-id="77bb4-256">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="77bb4-256">Office Add-ins XML manifest</span></span>](add-in-manifests.md)
- [<span data-ttu-id="77bb4-257">Office アドインの要件セット</span><span class="sxs-lookup"><span data-stu-id="77bb4-257">Office Add-in requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="77bb4-258">Word-Add-in-Get-Set-EditOpen-XML</span><span class="sxs-lookup"><span data-stu-id="77bb4-258">Word-Add-in-Get-Set-EditOpen-XML</span></span>](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML)
