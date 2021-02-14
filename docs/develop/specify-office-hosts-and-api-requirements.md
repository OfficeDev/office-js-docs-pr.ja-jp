---
title: Office のホストと API の要件を指定する
description: アドインが期待Office動作するために必要なアプリケーションと API の要件を指定する方法について説明します。
ms.date: 08/24/2020
localization_priority: Normal
ms.openlocfilehash: 948e86e99150ebf2d0bc7deaa5512627679b025f
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237841"
---
# <a name="specify-office-applications-and-api-requirements"></a><span data-ttu-id="c5588-103">Office アプリケーションと API 要件を指定する</span><span class="sxs-lookup"><span data-stu-id="c5588-103">Specify Office applications and API requirements</span></span>

<span data-ttu-id="c5588-104">アドインOffice期待通り動作するために、特定の Office アプリケーション、要件セット、API メンバー、または API のバージョンに依存している場合があります。</span><span class="sxs-lookup"><span data-stu-id="c5588-104">Your Office Add-in might depend on a specific Office application, a requirement set, an API member, or a version of the API in order to work as expected.</span></span> <span data-ttu-id="c5588-105">たとえば、次のようなアドインがあります。</span><span class="sxs-lookup"><span data-stu-id="c5588-105">For example, your add-in might:</span></span>

- <span data-ttu-id="c5588-106">1 つの Office アプリケーション (例: Word または Excel)、またはいくつかのアプリケーションで実行します。</span><span class="sxs-lookup"><span data-stu-id="c5588-106">Run in a single Office application (e.g., Word or Excel), or several applications.</span></span>

- <span data-ttu-id="c5588-p102">Office の一部のバージョンでのみ利用できる JavaScript API を使用します。たとえば、Excel 2016 で実行するアドインでは、Excel JavaScript API を使用することがあります。</span><span class="sxs-lookup"><span data-stu-id="c5588-p102">Make use of JavaScript APIs that are only available in some versions of Office. For example, you might use the Excel JavaScript APIs in an add-in that runs in Excel 2016.</span></span>

- <span data-ttu-id="c5588-109">アドインが使用する API メンバーをサポートするバージョンの Office でのみ実行します。</span><span class="sxs-lookup"><span data-stu-id="c5588-109">Run only in versions of Office that support API members that your add-in uses.</span></span>

<span data-ttu-id="c5588-110">この記事は、アドインが期待どおりに機能し、できるだけ多くのユーザーが利用できるようにするために選択する必要のあるオプションについて理解するのに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="c5588-110">This article helps you understand which options you should choose to ensure that your add-in works as expected and reaches the broadest audience possible.</span></span>

> [!NOTE]
> <span data-ttu-id="c5588-111">Office アドインが現在サポートされている場所の大きなレベルについては、「Office クライアント アプリケーションと Office[](../overview/office-add-in-availability.md)アドインのプラットフォームの可用性」ページを参照してください。</span><span class="sxs-lookup"><span data-stu-id="c5588-111">For a high-level view of where Office Add-ins are currently supported, see the [Office client application and platform availability for Office Add-ins](../overview/office-add-in-availability.md) page.</span></span>

<span data-ttu-id="c5588-112">この記事で説明する中心的な概念を次の表に示します。</span><span class="sxs-lookup"><span data-stu-id="c5588-112">The following table lists core concepts discussed throughout this article.</span></span>

|<span data-ttu-id="c5588-113">**概念**</span><span class="sxs-lookup"><span data-stu-id="c5588-113">**Concept**</span></span>|<span data-ttu-id="c5588-114">**説明**</span><span class="sxs-lookup"><span data-stu-id="c5588-114">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="c5588-115">Office、Office クライアント アプリケーション</span><span class="sxs-lookup"><span data-stu-id="c5588-115">Office application, Office client application</span></span>|<span data-ttu-id="c5588-p103">アドインの実行に使用される Office アプリケーション。たとえば、Word や Excel など。</span><span class="sxs-lookup"><span data-stu-id="c5588-p103">The Office application used to run your add-in. For example, Word, Excel, and so on.</span></span>|
|<span data-ttu-id="c5588-118">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="c5588-118">Platform</span></span>|<span data-ttu-id="c5588-119">ブラウザー Office iPad など、アプリケーションが実行される場所。</span><span class="sxs-lookup"><span data-stu-id="c5588-119">Where the Office application runs, such as in a browser or on an iPad.</span></span>|
|<span data-ttu-id="c5588-120">要件セット</span><span class="sxs-lookup"><span data-stu-id="c5588-120">Requirement set</span></span>|<span data-ttu-id="c5588-121">関連する API メンバーの名前付きグループ。</span><span class="sxs-lookup"><span data-stu-id="c5588-121">A named group of related API members.</span></span> <span data-ttu-id="c5588-122">アドインは要件セットを使用して、アドインが使用Office API メンバーをサポートするかどうかを判断します。</span><span class="sxs-lookup"><span data-stu-id="c5588-122">Add-ins use requirement sets to determine whether the Office application supports API members used by your add-in.</span></span> <span data-ttu-id="c5588-123">個々の API メンバーのサポートをテストするよりも、要件セットのサポートをテストするほうが簡単です。</span><span class="sxs-lookup"><span data-stu-id="c5588-123">It's easier to test for the support of a requirement set than for the support of individual API members.</span></span> <span data-ttu-id="c5588-124">要件セットのサポートは、OfficeアプリケーションとアプリケーションのバージョンによってOfficeされます。</span><span class="sxs-lookup"><span data-stu-id="c5588-124">Requirement set support varies by Office application and the version of the Office application.</span></span> <br ><span data-ttu-id="c5588-125">要件セットはマニフェスト ファイルで指定されます。</span><span class="sxs-lookup"><span data-stu-id="c5588-125">Requirement sets are specified in the manifest file.</span></span> <span data-ttu-id="c5588-126">マニフェストで要件セットを指定する場合は、アドインを実行するために Office アプリケーションが提供する必要がある API サポートの最小レベルを設定します。</span><span class="sxs-lookup"><span data-stu-id="c5588-126">When you specify requirement sets in the manifest, you set the minimum level of API support that the Office application must provide in order to run your add-in.</span></span> <span data-ttu-id="c5588-127">Officeで指定された要件セットをサポートしないアプリケーションはアドインを実行できないので、アドインは [マイ アドイン] に <span class="ui">表示されません</span>。これにより、アドインを使用できる場所が制限されます。</span><span class="sxs-lookup"><span data-stu-id="c5588-127">Office applications that don't support requirement sets specified in the manifest can't run your add-in, and your add-in won't display in <span class="ui">My Add-ins</span>. This restricts where your add-in is available.</span></span> <span data-ttu-id="c5588-128">コードでは、ランタイム チェックを使用します。</span><span class="sxs-lookup"><span data-stu-id="c5588-128">In code using runtime checks.</span></span> <span data-ttu-id="c5588-129">要件セットの詳細な一覧については、「[Office アドインの要件セット](../reference/requirement-sets/office-add-in-requirement-sets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c5588-129">For the complete list of requirement sets, see [Office Add-in requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md).</span></span>|
|<span data-ttu-id="c5588-130">ランタイム チェック</span><span class="sxs-lookup"><span data-stu-id="c5588-130">Runtime check</span></span>|<span data-ttu-id="c5588-131">アドインを実行している Office アプリケーションが、アドインで使用される要件セットまたはメソッドをサポートするかどうかを判断するために実行時に実行されるテスト。</span><span class="sxs-lookup"><span data-stu-id="c5588-131">A test that is performed at runtime to determine whether the Office application running your add-in supports requirement sets or methods used by your add-in.</span></span> <span data-ttu-id="c5588-132">ランタイム チェックを実行するには **、if** ステートメントをメソッド、要件セット、または要件セットの一部ではないメソッド名と一緒 `isSetSupported` に使用します。</span><span class="sxs-lookup"><span data-stu-id="c5588-132">To perform a runtime check, you use an **if** statement with the `isSetSupported` method, the requirement sets, or the method names that aren't part of a requirement set.</span></span> <span data-ttu-id="c5588-133">ランタイム チェックを使用すると、アドインを、最も多くのお客様が利用できるものにできます。</span><span class="sxs-lookup"><span data-stu-id="c5588-133">Use runtime checks to ensure that your add-in reaches the broadest number of customers.</span></span> <span data-ttu-id="c5588-134">要件セットとは異なり、ランタイム チェックでは、アドインを実行するために Office アプリケーションが提供する必要がある最小レベルの API サポートは指定されません。</span><span class="sxs-lookup"><span data-stu-id="c5588-134">Unlike requirement sets, runtime checks don't specify the minimum level of API support that the Office application must provide for your add-in to run.</span></span> <span data-ttu-id="c5588-135">代わりに **、if** ステートメントを使用して、API メンバーがサポートされているかどうかを判断します。</span><span class="sxs-lookup"><span data-stu-id="c5588-135">Instead, you use the **if** statement to determine whether an API member is supported.</span></span> <span data-ttu-id="c5588-136">サポートされている場合には、アドインで追加機能を提供できます。</span><span class="sxs-lookup"><span data-stu-id="c5588-136">If it is, you can provide additional functionality in your add-in.</span></span> <span data-ttu-id="c5588-137">ランタイム チェックを使用するときは、自分のアドインは必ず **[個人用アドイン]** に表示されます。</span><span class="sxs-lookup"><span data-stu-id="c5588-137">Your add-in will always display in **My Add-ins** when you use runtime checks.</span></span>|

## <a name="before-you-begin"></a><span data-ttu-id="c5588-138">始める前に</span><span class="sxs-lookup"><span data-stu-id="c5588-138">Before you begin</span></span>

<span data-ttu-id="c5588-139">アドインで最新バージョンのアドイン マニフェスト スキーマを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c5588-139">Your add-in must use the most current version of the add-in manifest schema.</span></span> <span data-ttu-id="c5588-140">アドインでランタイム チェックを使用する場合は、最新の JavaScript API (Office) ライブラリを使用office.jsしてください。</span><span class="sxs-lookup"><span data-stu-id="c5588-140">If you use runtime checks in your add-in, ensure that you use the latest Office JavaScript API (office.js) library.</span></span>

### <a name="specify-the-latest-add-in-manifest-schema"></a><span data-ttu-id="c5588-141">最新のアドイン マニフェスト スキーマを指定する</span><span class="sxs-lookup"><span data-stu-id="c5588-141">Specify the latest add-in manifest schema</span></span>

<span data-ttu-id="c5588-142">アドインのマニフェストでは、アドイン マニフェスト スキーマのバージョン 1.1 を使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c5588-142">Your add-in's manifest must use version 1.1 of the add-in manifest schema.</span></span> <span data-ttu-id="c5588-143">アドイン マニフェスト `OfficeApp` の要素を次のように設定します。</span><span class="sxs-lookup"><span data-stu-id="c5588-143">Set the `OfficeApp` element in your add-in manifest as follows.</span></span>

```XML
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
```

### <a name="specify-the-latest-office-javascript-api-library"></a><span data-ttu-id="c5588-144">JavaScript API ライブラリOffice指定する</span><span class="sxs-lookup"><span data-stu-id="c5588-144">Specify the latest Office JavaScript API library</span></span>

<span data-ttu-id="c5588-145">ランタイム チェックを使用する場合は、コンテンツ配信ネットワーク (CDN) から Office JavaScript API ライブラリの最新バージョンを参照します。</span><span class="sxs-lookup"><span data-stu-id="c5588-145">If you use runtime checks, reference the most current version of the Office JavaScript API library from the content delivery network (CDN).</span></span> <span data-ttu-id="c5588-146">その場合、HTML に次の `script` タグを追加します。</span><span class="sxs-lookup"><span data-stu-id="c5588-146">To do this, add the following  `script` tag to your HTML.</span></span> <span data-ttu-id="c5588-147">CDN URL で `/1/` を使用することで、Office.js の最新版が参照されます。</span><span class="sxs-lookup"><span data-stu-id="c5588-147">Using `/1/` in the CDN URL ensures that you reference the most recent version of Office.js.</span></span>

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

## <a name="options-to-specify-office-applications-or-api-requirements"></a><span data-ttu-id="c5588-148">アプリケーションまたは API Officeを指定するオプション</span><span class="sxs-lookup"><span data-stu-id="c5588-148">Options to specify Office applications or API requirements</span></span>

<span data-ttu-id="c5588-149">アプリケーションまたは API Officeを指定する場合は、いくつかの要素を考慮する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c5588-149">When you specify Office applications or API requirements, there are several factors to consider.</span></span> <span data-ttu-id="c5588-150">次の図に、アドインで使用すべき手法の判別方法を示します。</span><span class="sxs-lookup"><span data-stu-id="c5588-150">The following diagram shows how to decide which technique to use in your add-in.</span></span>

![アプリケーションまたは API の要件を指定するときに、アドインにOfficeオプションを選択する](../images/options-for-office-hosts.png)

- <span data-ttu-id="c5588-152">アドインが 1 つのアプリケーションで実行Officeマニフェスト `Hosts` で要素を設定します。</span><span class="sxs-lookup"><span data-stu-id="c5588-152">If your add-in runs in one Office application, set the `Hosts` element in the manifest.</span></span> <span data-ttu-id="c5588-153">詳しくは、「[Hosts 要素を設定する](#set-the-hosts-element)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c5588-153">For more information, see [Set the Hosts element](#set-the-hosts-element).</span></span>

- <span data-ttu-id="c5588-154">アドインを実行するために Office アプリケーションがサポートする必要がある最小要件セットまたは API メンバーを設定するには、マニフェストで要素 `Requirements` を設定します。</span><span class="sxs-lookup"><span data-stu-id="c5588-154">To set the minimum requirement set or API members that an Office application must support to run your add-in, set the `Requirements` element in the manifest.</span></span> <span data-ttu-id="c5588-155">詳しくは、「[マニフェストで Requirements 要素を設定する](#set-the-requirements-element-in-the-manifest)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="c5588-155">For more information, see [Set the Requirements element in the manifest](#set-the-requirements-element-in-the-manifest).</span></span>

- <span data-ttu-id="c5588-156">特定の要件セットまたは API メンバーが Office アプリケーションで使用可能な場合に追加の機能を提供する場合は、アドインの JavaScript コードでランタイム チェックを実行します。</span><span class="sxs-lookup"><span data-stu-id="c5588-156">If you would like to provide additional functionality if specific requirement sets or API members are available in the Office application, perform a runtime check in your add-in's JavaScript code.</span></span> <span data-ttu-id="c5588-157">たとえば、アドインが Excel 2016 で機能する場合は、Excel JavaScript API の API メンバーを使用して追加の機能を提供します。</span><span class="sxs-lookup"><span data-stu-id="c5588-157">For example, if your add-in runs in Excel 2016, use API members from the Excel JavaScript API to provide additional functionality.</span></span> <span data-ttu-id="c5588-158">詳細については、「[JavaScript コードでランタイム チェックを使用する](#use-runtime-checks-in-your-javascript-code)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="c5588-158">For more information, see [Use runtime checks in your JavaScript code](#use-runtime-checks-in-your-javascript-code).</span></span>

## <a name="set-the-hosts-element"></a><span data-ttu-id="c5588-159">Hosts 要素を設定する</span><span class="sxs-lookup"><span data-stu-id="c5588-159">Set the Hosts element</span></span>

<span data-ttu-id="c5588-160">アドインを 1 つのクライアント アプリケーションでOfficeするには、マニフェストの `Hosts` 要素 `Host` を使用します。</span><span class="sxs-lookup"><span data-stu-id="c5588-160">To make your add-in run in one Office client application, use the `Hosts` and `Host` elements in the manifest.</span></span> <span data-ttu-id="c5588-161">要素を指定しない場合、アドインは、アドインによってサポートOfficeのすべての Office アプリケーション `Hosts` で実行されます。</span><span class="sxs-lookup"><span data-stu-id="c5588-161">If you don't specify the `Hosts` element, your add-in will run in all Office applications supported by Office Add-ins.</span></span>

<span data-ttu-id="c5588-162">たとえば、次と宣言は、アドインが Excel のリリース `Hosts` (Excel on the web、Windows、iPad など) で動作する場合に指定 `Host` します。</span><span class="sxs-lookup"><span data-stu-id="c5588-162">For example, the following `Hosts` and `Host` declaration specifies that the add-in will work with any release of Excel, which includes Excel on the web, Windows, and iPad.</span></span>

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

<span data-ttu-id="c5588-163">要素 `Hosts` には、1 つ以上の要素を含 `Host` めできます。</span><span class="sxs-lookup"><span data-stu-id="c5588-163">The `Hosts` element can contain one or more `Host` elements.</span></span> <span data-ttu-id="c5588-164">この `Host` 要素は、アドインOffice必要なアプリケーションを指定します。</span><span class="sxs-lookup"><span data-stu-id="c5588-164">The `Host` element specifies the Office application your add-in requires.</span></span> <span data-ttu-id="c5588-165">この `Name` 属性は必須であり、次のいずれかの値に設定できます。</span><span class="sxs-lookup"><span data-stu-id="c5588-165">The `Name` attribute is required and can be set to one of the following values.</span></span>

| <span data-ttu-id="c5588-166">名前</span><span class="sxs-lookup"><span data-stu-id="c5588-166">Name</span></span>          | <span data-ttu-id="c5588-167">Office クライアント アプリケーション</span><span class="sxs-lookup"><span data-stu-id="c5588-167">Office client applications</span></span>                      |
|:--------------|:----------------------------------------------|
| <span data-ttu-id="c5588-168">データベース</span><span class="sxs-lookup"><span data-stu-id="c5588-168">Database</span></span>      | <span data-ttu-id="c5588-169">Access Web アプリ</span><span class="sxs-lookup"><span data-stu-id="c5588-169">Access web apps</span></span>                               |
| <span data-ttu-id="c5588-170">Document</span><span class="sxs-lookup"><span data-stu-id="c5588-170">Document</span></span>      | <span data-ttu-id="c5588-171">Word on the web、Windows、Mac、iPad</span><span class="sxs-lookup"><span data-stu-id="c5588-171">Word on the web, Windows, Mac, iPad</span></span>           |
| <span data-ttu-id="c5588-172">メールボックス</span><span class="sxs-lookup"><span data-stu-id="c5588-172">Mailbox</span></span>       | <span data-ttu-id="c5588-173">Outlook on the web、Windows、Mac、Android、iOS</span><span class="sxs-lookup"><span data-stu-id="c5588-173">Outlook on the web, Windows, Mac, Android, iOS</span></span>|
| <span data-ttu-id="c5588-174">Presentation</span><span class="sxs-lookup"><span data-stu-id="c5588-174">Presentation</span></span>  | <span data-ttu-id="c5588-175">PowerPoint on the web、Windows、Mac、iPad</span><span class="sxs-lookup"><span data-stu-id="c5588-175">PowerPoint on the web, Windows, Mac, iPad</span></span>     |
| <span data-ttu-id="c5588-176">Project</span><span class="sxs-lookup"><span data-stu-id="c5588-176">Project</span></span>       | <span data-ttu-id="c5588-177">Windows での Project</span><span class="sxs-lookup"><span data-stu-id="c5588-177">Project on Windows</span></span>                            |
| <span data-ttu-id="c5588-178">Workbook</span><span class="sxs-lookup"><span data-stu-id="c5588-178">Workbook</span></span>      | <span data-ttu-id="c5588-179">Excel on the web、Windows、Mac、iPad</span><span class="sxs-lookup"><span data-stu-id="c5588-179">Excel on the web, Windows, Mac, iPad</span></span>          |

> [!NOTE]
> <span data-ttu-id="c5588-180">この `Name` 属性は、アドインOffice実行できるクライアント アプリケーションの名前を指定します。</span><span class="sxs-lookup"><span data-stu-id="c5588-180">The `Name` attribute specifies the Office client application that can run your add-in.</span></span> <span data-ttu-id="c5588-181">Officeアプリケーションは、さまざまなプラットフォームでサポートされ、デスクトップ、Web ブラウザー、タブレット、モバイル デバイスで実行されます。</span><span class="sxs-lookup"><span data-stu-id="c5588-181">Office applications are supported on different platforms and run on desktops, web browsers, tablets, and mobile devices.</span></span> <span data-ttu-id="c5588-182">アドインを実行するために使用するプラットフォームを指定することはできません。</span><span class="sxs-lookup"><span data-stu-id="c5588-182">You can't specify which platform can be used to run your add-in.</span></span> <span data-ttu-id="c5588-183">たとえば、指定した場合、Outlook on the web と Windows の両方を使用してアドイン `Mailbox` を実行できます。</span><span class="sxs-lookup"><span data-stu-id="c5588-183">For example, if you specify `Mailbox`, both Outlook on the web and on Windows can be used to run your add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c5588-184">SharePoint で Access Web アプリとデータベースを作成して使用することは推奨されなくなりました。</span><span class="sxs-lookup"><span data-stu-id="c5588-184">We no longer recommend that you create and use Access web apps and databases in SharePoint.</span></span> <span data-ttu-id="c5588-185">代わりに、[Microsoft PowerApps](https://powerapps.microsoft.com/) を使用して、コード作成が不要な Web とモバイル デバイス用ビジネス ソリューションをビルドすることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="c5588-185">As an alternative, we recommend that you use [Microsoft PowerApps](https://powerapps.microsoft.com/) to build no-code business solutions for web and mobile devices.</span></span>

## <a name="set-the-requirements-element-in-the-manifest"></a><span data-ttu-id="c5588-186">マニフェストで Requirements 要素を設定する</span><span class="sxs-lookup"><span data-stu-id="c5588-186">Set the Requirements element in the manifest</span></span>

<span data-ttu-id="c5588-187">この要素は、アドインを実行するために、Office アプリケーションでサポートする必要がある最小要件セットまたは API メンバー `Requirements` を指定します。</span><span class="sxs-lookup"><span data-stu-id="c5588-187">The `Requirements` element specifies the minimum requirement sets or API members that must be supported by the Office application to run your add-in.</span></span> <span data-ttu-id="c5588-188">この `Requirements` 要素は、要件セットと、アドインで使用される個々のメソッドの両方を指定できます。</span><span class="sxs-lookup"><span data-stu-id="c5588-188">The `Requirements` element can specify both requirement sets and individual methods used in your add-in.</span></span> <span data-ttu-id="c5588-189">アドイン マニフェスト スキーマのバージョン 1.1 では、Outlook アドインを除くすべてのアドインでこの要素は `Requirements` オプションです。</span><span class="sxs-lookup"><span data-stu-id="c5588-189">In version 1.1 of the add-in manifest schema, the `Requirements` element is optional for all add-ins, except for Outlook add-ins.</span></span>

> [!WARNING]
> <span data-ttu-id="c5588-190">この要素は `Requirements` 、アドインで使用する必要がある重要な要件セットまたは API メンバーを指定する場合にのみ使用します。</span><span class="sxs-lookup"><span data-stu-id="c5588-190">Only use the `Requirements` element to specify critical requirement sets or API members that your add-in must use.</span></span> <span data-ttu-id="c5588-191">Office アプリケーションまたはプラットフォームが要素で指定された要件セットまたは API メンバーをサポートしない場合、アドインは、そのアプリケーションまたはプラットフォームでは実行されません。また、マイ アドインには表示されません `Requirements` 。 代わりに、Excel on the web、Windows、iPad など、Office アプリケーションのすべてのプラットフォームでアドインを使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="c5588-191">If the Office application or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that application or platform, and won't display in **My Add-ins**. Instead, we recommend that you make your add-in available on all platforms of an Office application, such as Excel on the web, Windows, and iPad.</span></span> <span data-ttu-id="c5588-192">アドインをすべてのアプリケーションとプラットフォーム  _でOfficeするには_ 、要素の代わりにランタイム チェックを使用 `Requirements` します。</span><span class="sxs-lookup"><span data-stu-id="c5588-192">To make your add-in available on  _all_ Office applications and platforms, use runtime checks instead of the `Requirements` element.</span></span>

<span data-ttu-id="c5588-193">次のコード例は、以下をサポートしているすべてのクライアント Office読み込まれるアドインを示しています。</span><span class="sxs-lookup"><span data-stu-id="c5588-193">The following code example shows an add-in that loads in all Office client applications that support the following:</span></span>

-  <span data-ttu-id="c5588-194">`TableBindings` 要件セット。最小バージョンは "1.1" です。</span><span class="sxs-lookup"><span data-stu-id="c5588-194">`TableBindings` requirement set, which has a minimum version of "1.1".</span></span>

-  <span data-ttu-id="c5588-195">`OOXML` 要件セット。最小バージョンは "1.1" です。</span><span class="sxs-lookup"><span data-stu-id="c5588-195">`OOXML` requirement set, which has a minimum version of "1.1".</span></span>

-  <span data-ttu-id="c5588-196">`Document.getSelectedDataAsync` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="c5588-196">`Document.getSelectedDataAsync` method.</span></span>

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

- <span data-ttu-id="c5588-197">要素 `Requirements` には、子要素 `Sets` `Methods` と子要素が含まれます。</span><span class="sxs-lookup"><span data-stu-id="c5588-197">The `Requirements` element contains the `Sets` and `Methods` child elements.</span></span>

- <span data-ttu-id="c5588-198">要素 `Sets` には、1 つ以上の要素を含 `Set` めできます。</span><span class="sxs-lookup"><span data-stu-id="c5588-198">The `Sets` element can contain one or more `Set` elements.</span></span> <span data-ttu-id="c5588-199">`DefaultMinVersion` すべての子要素の `MinVersion` 既定値を指定 `Set` します。</span><span class="sxs-lookup"><span data-stu-id="c5588-199">`DefaultMinVersion` specifies the default `MinVersion` value of all child `Set` elements.</span></span>

- <span data-ttu-id="c5588-200">この `Set` 要素は、アドインを実行するためにOfficeアプリケーションがサポートする必要がある要件セットを指定します。</span><span class="sxs-lookup"><span data-stu-id="c5588-200">The `Set` element specifies requirement sets that the Office application must support to run the add-in.</span></span> <span data-ttu-id="c5588-201">この `Name` 属性は、要件セットの名前を指定します。</span><span class="sxs-lookup"><span data-stu-id="c5588-201">The `Name` attribute specifies the name of the requirement set.</span></span> <span data-ttu-id="c5588-202">要件 `MinVersion` セットの最小バージョンを指定します。</span><span class="sxs-lookup"><span data-stu-id="c5588-202">The `MinVersion` specifies the minimum version of the requirement set.</span></span> <span data-ttu-id="c5588-203">`MinVersion` overrides the value of `DefaultMinVersion` Requirement sets and requirement set versions that your API members belong to, see Office [Add-in requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="c5588-203">`MinVersion` overrides the value of `DefaultMinVersion` For more information about requirement sets and requirement set versions that your API members belong to, see [Office Add-in requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md).</span></span>

- <span data-ttu-id="c5588-p122">要素 `Methods` には、1 つ以上の要素を含 `Method` めできます。この要素を Outlook `Methods` アドインと一緒に使用することはできません。</span><span class="sxs-lookup"><span data-stu-id="c5588-p122">The `Methods` element can contain one or more `Method` elements. You can't use the `Methods` element with Outlook add-ins.</span></span>

- <span data-ttu-id="c5588-p123">この `Method` 要素は、アドインを実行するアプリケーションでサポートOfficeする必要がある個別のメソッドを指定します。この `Name` 属性は必須であり、親オブジェクトで修飾されたメソッドの名前を指定します。</span><span class="sxs-lookup"><span data-stu-id="c5588-p123">The `Method` element specifies an individual method that must be supported in the Office application where your add-in runs. The `Name` attribute is required and specifies the name of the method qualified with its parent object.</span></span>

## <a name="use-runtime-checks-in-your-javascript-code"></a><span data-ttu-id="c5588-208">JavaScript コードでランタイム チェックを使用する</span><span class="sxs-lookup"><span data-stu-id="c5588-208">Use runtime checks in your JavaScript code</span></span>

<span data-ttu-id="c5588-209">特定の要件セットがアプリケーションでサポートされている場合は、アドインに追加の機能をOfficeがあります。</span><span class="sxs-lookup"><span data-stu-id="c5588-209">You might want to provide additional functionality in your add-in if certain requirement sets are supported by the Office application.</span></span> <span data-ttu-id="c5588-210">たとえば、アドインで Word 2016 を実行する場合、既存のアドインで Word JavaScript API を使用することがあります。</span><span class="sxs-lookup"><span data-stu-id="c5588-210">For example, you might want to use the Word JavaScript APIs in your existing add-in if your add-in runs in Word 2016.</span></span> <span data-ttu-id="c5588-211">その場合、要件セットの名前を指定し、[isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="c5588-211">To do this, you use the [isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) method with the name of the requirement set.</span></span> <span data-ttu-id="c5588-212">`isSetSupported` 実行時に、アドインを実行Officeアプリケーションが要件セットをサポートするかどうかを決定します。</span><span class="sxs-lookup"><span data-stu-id="c5588-212">`isSetSupported` determines, at runtime, whether the Office application running the add-in supports the requirement set.</span></span> <span data-ttu-id="c5588-213">要件セットがサポートされている場合は true を返し、その要件セットの API メンバーを使用する追加コード `isSetSupported` を実行します。 </span><span class="sxs-lookup"><span data-stu-id="c5588-213">If the requirement set is supported, `isSetSupported` returns **true** and runs the additional code that uses the API members from that requirement set.</span></span> <span data-ttu-id="c5588-214">アプリケーションがOfficeが要件セットをサポートしない場合は false を返し、追加のコード `isSetSupported` は実行されません。 </span><span class="sxs-lookup"><span data-stu-id="c5588-214">If the Office application doesn't support the requirement set, `isSetSupported` returns **false** and the additional code won't run.</span></span> <span data-ttu-id="c5588-215">次のコードは `isSetSupported` と共に使用する構文を示しています。</span><span class="sxs-lookup"><span data-stu-id="c5588-215">The following code shows the syntax to use with `isSetSupported`.</span></span>

```js
if (Office.context.requirements.isSetSupported(RequirementSetName, MinimumVersion))
{
   // Code that uses API members from RequirementSetName.
}

```

- <span data-ttu-id="c5588-216">_RequirementSetName_ (必須) は、要件セットの名前を表す文字列です (例: "**ExcelApi**"、"**Mailbox**" など)。</span><span class="sxs-lookup"><span data-stu-id="c5588-216">_RequirementSetName_ (required) is a string that represents the name of the requirement set (e.g., "**ExcelApi**", "**Mailbox**", etc.).</span></span> <span data-ttu-id="c5588-217">利用できる要件セットの詳細については、「[Office アドインの要件セット](../reference/requirement-sets/office-add-in-requirement-sets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c5588-217">For more information about available requirement sets, see [Office Add-in requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md).</span></span>
- <span data-ttu-id="c5588-218">_MinimumVersion_ (オプション) は、ステートメント内のコードを実行するために Office アプリケーションがサポートする必要がある最小要件セットのバージョンを指定する文字列です (例: `if` "**1.9**")。</span><span class="sxs-lookup"><span data-stu-id="c5588-218">_MinimumVersion_ (optional) is a string that specifies the minimum requirement set version that the Office application must support in order for the code within the `if` statement to run (e.g., "**1.9**").</span></span>

> [!WARNING]
> <span data-ttu-id="c5588-219">メソッドを呼 `isSetSupported` び出す場合、パラメーターの値 (指定されている場合 `MinimumVersion` ) は文字列である必要があります。</span><span class="sxs-lookup"><span data-stu-id="c5588-219">When calling the `isSetSupported` method, the value of the `MinimumVersion` parameter (if specified) should be a string.</span></span> <span data-ttu-id="c5588-220">これは、JavaScript パーサーでは、1.1 や 1.10 のような数値の間の差異を区別できないが、"1.1" や "1.10" などの文字列値ではできるからです。</span><span class="sxs-lookup"><span data-stu-id="c5588-220">This is because the JavaScript parser cannot differentiate between numeric values such as 1.1 and 1.10, where as it can for string values such as "1.1" and "1.10".</span></span>
> <span data-ttu-id="c5588-221">`number` のオーバーロードは非推奨になります。</span><span class="sxs-lookup"><span data-stu-id="c5588-221">The `number` overload is deprecated.</span></span>

<span data-ttu-id="c5588-222">次 `isSetSupported` のように、 `RequirementSetName` アプリケーションに関連付Office使用します。</span><span class="sxs-lookup"><span data-stu-id="c5588-222">Use `isSetSupported` with the `RequirementSetName` associated with the Office application as follows.</span></span>

|<span data-ttu-id="c5588-223">Office アプリケーション</span><span class="sxs-lookup"><span data-stu-id="c5588-223">Office application</span></span>|<span data-ttu-id="c5588-224">RequirementSetName</span><span class="sxs-lookup"><span data-stu-id="c5588-224">RequirementSetName</span></span>|
|---|---|
|<span data-ttu-id="c5588-225">Excel</span><span class="sxs-lookup"><span data-stu-id="c5588-225">Excel</span></span>|<span data-ttu-id="c5588-226">ExcelApi</span><span class="sxs-lookup"><span data-stu-id="c5588-226">ExcelApi</span></span>|
|<span data-ttu-id="c5588-227">OneNote</span><span class="sxs-lookup"><span data-stu-id="c5588-227">OneNote</span></span>|<span data-ttu-id="c5588-228">OneNoteApi</span><span class="sxs-lookup"><span data-stu-id="c5588-228">OneNoteApi</span></span>|
|<span data-ttu-id="c5588-229">Outlook</span><span class="sxs-lookup"><span data-stu-id="c5588-229">Outlook</span></span>|<span data-ttu-id="c5588-230">Mailbox</span><span class="sxs-lookup"><span data-stu-id="c5588-230">Mailbox</span></span>|
|<span data-ttu-id="c5588-231">Word</span><span class="sxs-lookup"><span data-stu-id="c5588-231">Word</span></span>|<span data-ttu-id="c5588-232">WordApi</span><span class="sxs-lookup"><span data-stu-id="c5588-232">WordApi</span></span>|

<span data-ttu-id="c5588-233">これらの `isSetSupported` アプリケーションのメソッドと要件セットは、CDN の最新の Office.js ファイルで使用できます。</span><span class="sxs-lookup"><span data-stu-id="c5588-233">The `isSetSupported` method and the requirement sets for these applications are available in the latest Office.js file on the CDN.</span></span> <span data-ttu-id="c5588-234">CDN からのカスタム Office.js使用しない場合は、未定義のため、アドインで例外 `isSetSupported` が生成される可能性があります。</span><span class="sxs-lookup"><span data-stu-id="c5588-234">If you don't use Office.js from the CDN, your add-in might generate exceptions because `isSetSupported` will be undefined.</span></span> <span data-ttu-id="c5588-235">詳細については [、「JavaScript API ライブラリの最新のOffice指定する」を参照してください](#specify-the-latest-office-javascript-api-library)。</span><span class="sxs-lookup"><span data-stu-id="c5588-235">For more information, see [Specify the latest Office JavaScript API library](#specify-the-latest-office-javascript-api-library).</span></span>

<span data-ttu-id="c5588-236">次のコード例は、さまざまな要件セットまたは API メンバーをサポートする可能性があるさまざまな Office アプリケーションに対して、アドインがさまざまな機能を提供する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="c5588-236">The following code example shows how an add-in can provide different functionality for different Office applications that might support different requirement sets or API members.</span></span>

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

## <a name="runtime-checks-using-methods-not-in-a-requirement-set"></a><span data-ttu-id="c5588-237">要件セットにないメソッドを使用したランタイム チェック</span><span class="sxs-lookup"><span data-stu-id="c5588-237">Runtime checks using methods not in a requirement set</span></span>

<span data-ttu-id="c5588-238">API の一部のメンバーは、要件のセットに属していません。</span><span class="sxs-lookup"><span data-stu-id="c5588-238">Some API members don't belong to requirement sets.</span></span> <span data-ttu-id="c5588-239">[これは、Office JavaScript API](../reference/javascript-api-for-office.md)名前空間の一部である API メンバー (Outlook メールボックス API を除くすべてのメンバー) にのみ適用されますが、Word JavaScript API に属する API メンバー (次の場合は何でも含む)、Excel JavaScript API (何でも含む `Office.` [](/javascript/api/outlook)[](../reference/overview/word-add-ins-reference-overview.md) `Word.` [](../reference/overview/excel-add-ins-reference-overview.md) `Excel.` )、OneNote [JavaScript API](../reference/overview/onenote-add-ins-javascript-reference.md) (すべての名前空間 `OneNote.` ) には適用されません。</span><span class="sxs-lookup"><span data-stu-id="c5588-239">This only applies to API members that are part of the [Office JavaScript API](../reference/javascript-api-for-office.md) namespace (anything under `Office.` except [Outlook Mailbox APIs](/javascript/api/outlook)), but not API members that belong to the [Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) (anything in `Word.`), [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) (anything in `Excel.`), or [OneNote JavaScript API](../reference/overview/onenote-add-ins-javascript-reference.md) (anything in `OneNote.`) namespaces.</span></span> <span data-ttu-id="c5588-240">アドインが要件セットの一部ではないメソッドに依存している場合は、ランタイム チェックを使用して、次のコード例に示すように、メソッドが Office アプリケーションでサポートされているかどうかを判断できます。</span><span class="sxs-lookup"><span data-stu-id="c5588-240">When your add-in depends on a method that is not part of a requirement set, you can use the runtime check to determine whether the method is supported by the Office application, as shown in the following code example.</span></span> <span data-ttu-id="c5588-241">要件セットに属さないメソッドの詳細な一覧については、「[Office アドインの要件セット](../reference/requirement-sets/office-add-in-requirement-sets.md#methods-that-arent-part-of-a-requirement-set)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c5588-241">For a complete list of methods that don't belong to a requirement set, see [Office Add-in requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md#methods-that-arent-part-of-a-requirement-set).</span></span>

> [!NOTE]
> <span data-ttu-id="c5588-242">アドインのコードでのこの種のランタイム チェックは、限定的に使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="c5588-242">We recommend that you limit the use of this type of runtime check in your add-in's code.</span></span>

<span data-ttu-id="c5588-243">次のコード例では、アプリケーションがサポートOfficeチェックします `document.setSelectedDataAsync` 。</span><span class="sxs-lookup"><span data-stu-id="c5588-243">The following code example checks whether the Office application supports `document.setSelectedDataAsync`.</span></span>

```js
if (Office.context.document.setSelectedDataAsync)
{
    // Run code that uses `document.setSelectedDataAsync`.
}
```


## <a name="see-also"></a><span data-ttu-id="c5588-244">関連項目</span><span class="sxs-lookup"><span data-stu-id="c5588-244">See also</span></span>

- [<span data-ttu-id="c5588-245">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="c5588-245">Office Add-ins XML manifest</span></span>](add-in-manifests.md)
- [<span data-ttu-id="c5588-246">Office アドインの要件セット</span><span class="sxs-lookup"><span data-stu-id="c5588-246">Office Add-in requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="c5588-247">Word-Add-in-Get-Set-EditOpen-XML</span><span class="sxs-lookup"><span data-stu-id="c5588-247">Word-Add-in-Get-Set-EditOpen-XML</span></span>](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML)
