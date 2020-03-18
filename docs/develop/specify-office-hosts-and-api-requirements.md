---
title: Office のホストと API の要件を指定する
description: アドインが意図したとおりに動作するように Office のホストと API の要件を指定する方法について説明します。
ms.date: 09/26/2019
localization_priority: Normal
ms.openlocfilehash: ab9b97f3d3232339010179097e1fd03dbeb86aa2
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718812"
---
# <a name="specify-office-hosts-and-api-requirements"></a><span data-ttu-id="281d6-103">Office のホストと API の要件を指定する</span><span class="sxs-lookup"><span data-stu-id="281d6-103">Specify Office hosts and API requirements</span></span>

<span data-ttu-id="281d6-p101">期待どおりの動作をするうえで、Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。たとえば、次のようなアドインがあります。</span><span class="sxs-lookup"><span data-stu-id="281d6-p101">Your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API in order to work as expected. For example, your add-in might:</span></span>

- <span data-ttu-id="281d6-106">1 つの Office アプリケーション (例: Word または Excel)、またはいくつかのアプリケーションで実行します。</span><span class="sxs-lookup"><span data-stu-id="281d6-106">Run in a single Office application (e.g., Word or Excel), or several applications.</span></span>

- <span data-ttu-id="281d6-p102">Office の一部のバージョンでのみ利用できる JavaScript API を使用します。たとえば、Excel 2016 で実行するアドインでは、Excel JavaScript API を使用することがあります。</span><span class="sxs-lookup"><span data-stu-id="281d6-p102">Make use of JavaScript APIs that are only available in some versions of Office. For example, you might use the Excel JavaScript APIs in an add-in that runs in Excel 2016.</span></span>

- <span data-ttu-id="281d6-109">アドインが使用する API メンバーをサポートするバージョンの Office でのみ実行します。</span><span class="sxs-lookup"><span data-stu-id="281d6-109">Run only in versions of Office that support API members that your add-in uses.</span></span>

<span data-ttu-id="281d6-110">この記事は、アドインが期待どおりに機能し、できるだけ多くのユーザーが利用できるようにするために選択する必要のあるオプションについて理解するのに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="281d6-110">This article helps you understand which options you should choose to ensure that your add-in works as expected and reaches the broadest audience possible.</span></span>

> [!NOTE]
> <span data-ttu-id="281d6-111">現時点での Office アドインのサポート状況の概要については、「[Office アドインを使用できるホストおよびプラットフォーム](../overview/office-add-in-availability.md)」のページを参照してください。</span><span class="sxs-lookup"><span data-stu-id="281d6-111">For a high-level view of where Office Add-ins are currently supported, see the [Office Add-in host and platform availability](../overview/office-add-in-availability.md) page.</span></span>

<span data-ttu-id="281d6-112">この記事で説明する中心的な概念を次の表に示します。</span><span class="sxs-lookup"><span data-stu-id="281d6-112">The following table lists core concepts discussed throughout this article.</span></span>

|<span data-ttu-id="281d6-113">**概念**</span><span class="sxs-lookup"><span data-stu-id="281d6-113">**Concept**</span></span>|<span data-ttu-id="281d6-114">**説明**</span><span class="sxs-lookup"><span data-stu-id="281d6-114">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="281d6-115">Office アプリケーション、Office ホスト アプリケーション、Office ホスト、またはホスト</span><span class="sxs-lookup"><span data-stu-id="281d6-115">Office application, Office host application, Office host, or host</span></span>|<span data-ttu-id="281d6-p103">アドインの実行に使用される Office アプリケーション。たとえば、Word や Excel など。</span><span class="sxs-lookup"><span data-stu-id="281d6-p103">The Office application used to run your add-in. For example, Word, Excel, and so on.</span></span>|
|<span data-ttu-id="281d6-118">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="281d6-118">Platform</span></span>|<span data-ttu-id="281d6-119">Office ホストを実行する場所。ブラウザーや iPad など。</span><span class="sxs-lookup"><span data-stu-id="281d6-119">Where the Office host runs, such as in a browser or on an iPad.</span></span>|
|<span data-ttu-id="281d6-120">要件セット</span><span class="sxs-lookup"><span data-stu-id="281d6-120">Requirement set</span></span>|<span data-ttu-id="281d6-p104">関連する API メンバーの名前付きグループ。アドインは要件セットを使用して、Office ホストが、アドインによって使用される API メンバーをサポートしているかどうかを判別します。個々の API メンバーのサポートをテストするよりも、要件セットのサポートをテストするほうが簡単です。要件セットのサポートは、Office ホストと Office ホストのバージョンによって異なります。 </span><span class="sxs-lookup"><span data-stu-id="281d6-p104">A named group of related API members. Add-ins use requirement sets to determine whether the Office host supports API members used by your add-in. It's easier to test for the support of a requirement set than for the support of individual API members. Requirement set support varies by Office host and the version of the Office host. </span></span><br ><span data-ttu-id="281d6-125">要件セットはマニフェスト ファイルで指定されます。</span><span class="sxs-lookup"><span data-stu-id="281d6-125">Requirement sets are specified in the manifest file.</span></span> <span data-ttu-id="281d6-126">マニフェストで要件セットを指定するときは、アドインを実行するために Office ホストが提供する必要のある最小レベルの API サポートを設定します。</span><span class="sxs-lookup"><span data-stu-id="281d6-126">When you specify requirement sets in the manifest, you set the minimum level of API support that the Office host must provide in order to run your add-in.</span></span> <span data-ttu-id="281d6-127">マニフェストで指定されている要件セットをサポートしていない Office ホストはアドインを実行できず、アドインは <span class="ui">[個人用アドイン]</span> に表示されません。これにより、アドインが利用できる場所が制限されます。</span><span class="sxs-lookup"><span data-stu-id="281d6-127">Office hosts that don't support requirement sets specified in the manifest can't run your add-in, and your add-in won't display in <span class="ui">My Add-ins</span>. This restricts where your add-in is available.</span></span> <span data-ttu-id="281d6-128">コードでは、ランタイム チェックを使用します。</span><span class="sxs-lookup"><span data-stu-id="281d6-128">In code using runtime checks.</span></span> <span data-ttu-id="281d6-129">要件セットの詳細な一覧については、「[Office アドインの要件セット](../reference/requirement-sets/office-add-in-requirement-sets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="281d6-129">For the complete list of requirement sets, see [Office Add-in requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md).</span></span>|
|<span data-ttu-id="281d6-130">ランタイム チェック</span><span class="sxs-lookup"><span data-stu-id="281d6-130">Runtime check</span></span>|<span data-ttu-id="281d6-131">アドインを実行している Office ホストが、アドインで使用されている要件セットまたはメソッドをサポートしているかどうかを判別するために実行時に行われるテスト。</span><span class="sxs-lookup"><span data-stu-id="281d6-131">A test that is performed at runtime to determine whether the Office host running your add-in supports requirement sets or methods used by your add-in.</span></span> <span data-ttu-id="281d6-132">ランタイムチェックを実行するには、 `isSetSupported`メソッドの**if**ステートメント、要件セット、または要件セットの一部ではないメソッド名を使用します。</span><span class="sxs-lookup"><span data-stu-id="281d6-132">To perform a runtime check, you use an **if** statement with the `isSetSupported` method, the requirement sets, or the method names that aren't part of a requirement set.</span></span> <span data-ttu-id="281d6-133">ランタイム チェックを使用すると、アドインを、最も多くのお客様が利用できるものにできます。</span><span class="sxs-lookup"><span data-stu-id="281d6-133">Use runtime checks to ensure that your add-in reaches the broadest number of customers.</span></span> <span data-ttu-id="281d6-134">要件セットとは異なり、ランタイム チェックでは、対象アドインを実行するために Office ホストが提供する必要のある最小レベルの API サポートは指定しません。</span><span class="sxs-lookup"><span data-stu-id="281d6-134">Unlike requirement sets, runtime checks don't specify the minimum level of API support that the Office host must provide for your add-in to run.</span></span> <span data-ttu-id="281d6-135">代わりに、 **if**ステートメントを使用して、API メンバーがサポートされているかどうかを判断します。</span><span class="sxs-lookup"><span data-stu-id="281d6-135">Instead, you use the **if** statement to determine whether an API member is supported.</span></span> <span data-ttu-id="281d6-136">サポートされている場合には、アドインで追加機能を提供できます。</span><span class="sxs-lookup"><span data-stu-id="281d6-136">If it is, you can provide additional functionality in your add-in.</span></span> <span data-ttu-id="281d6-137">ランタイム チェックを使用するときは、自分のアドインは必ず **[個人用アドイン]** に表示されます。</span><span class="sxs-lookup"><span data-stu-id="281d6-137">Your add-in will always display in **My Add-ins** when you use runtime checks.</span></span>|

## <a name="before-you-begin"></a><span data-ttu-id="281d6-138">始める前に</span><span class="sxs-lookup"><span data-stu-id="281d6-138">Before you begin</span></span>

<span data-ttu-id="281d6-139">アドインで最新バージョンのアドイン マニフェスト スキーマを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="281d6-139">Your add-in must use the most current version of the add-in manifest schema.</span></span> <span data-ttu-id="281d6-140">アドインでランタイムチェックを使用する場合は、最新の Office JavaScript API (office .js) ライブラリを使用していることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="281d6-140">If you use runtime checks in your add-in, ensure that you use the latest Office JavaScript API (office.js) library.</span></span>

### <a name="specify-the-latest-add-in-manifest-schema"></a><span data-ttu-id="281d6-141">最新のアドイン マニフェスト スキーマを指定する</span><span class="sxs-lookup"><span data-stu-id="281d6-141">Specify the latest add-in manifest schema</span></span>

<span data-ttu-id="281d6-142">アドインのマニフェストでは、アドイン マニフェスト スキーマのバージョン 1.1 を使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="281d6-142">Your add-in's manifest must use version 1.1 of the add-in manifest schema.</span></span> <span data-ttu-id="281d6-143">アドインマニフェスト`OfficeApp`の要素を次のように設定します。</span><span class="sxs-lookup"><span data-stu-id="281d6-143">Set the `OfficeApp` element in your add-in manifest as follows.</span></span>

```XML
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
```

### <a name="specify-the-latest-office-javascript-api-library"></a><span data-ttu-id="281d6-144">最新の Office JavaScript API ライブラリを指定する</span><span class="sxs-lookup"><span data-stu-id="281d6-144">Specify the latest Office JavaScript API library</span></span>

<span data-ttu-id="281d6-145">ランタイムチェックを使用する場合は、コンテンツ配信ネットワーク (CDN) から、最新バージョンの Office JavaScript API ライブラリを参照します。</span><span class="sxs-lookup"><span data-stu-id="281d6-145">If you use runtime checks, reference the most current version of the Office JavaScript API library from the content delivery network (CDN).</span></span> <span data-ttu-id="281d6-146">その場合、HTML に次の `script` タグを追加します。</span><span class="sxs-lookup"><span data-stu-id="281d6-146">To do this, add the following  `script` tag to your HTML.</span></span> <span data-ttu-id="281d6-147">CDN URL で `/1/` を使用することで、Office.js の最新版が参照されます。</span><span class="sxs-lookup"><span data-stu-id="281d6-147">Using `/1/` in the CDN URL ensures that you reference the most recent version of Office.js.</span></span>

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

## <a name="options-to-specify-office-hosts-or-api-requirements"></a><span data-ttu-id="281d6-148">Office のホストや API の要件を指定するオプション</span><span class="sxs-lookup"><span data-stu-id="281d6-148">Options to specify Office hosts or API requirements</span></span>

<span data-ttu-id="281d6-p110">Office ホストまたは API の要件を指定するときに、検討すべき事項がいくつかあります。次の図に、アドインで使用すべき手法の判別方法を示します。</span><span class="sxs-lookup"><span data-stu-id="281d6-p110">When you specify Office hosts or API requirements, there are several factors to consider. The following diagram shows how to decide which technique to use in your add-in.</span></span>

![Office のホストまたは API の要件を指定する際に、アドインに最適なオプションを選択する](../images/options-for-office-hosts.png)

- <span data-ttu-id="281d6-152">アドインが1つの Office ホストで実行される`Hosts`場合は、マニフェスト内の要素を設定します。</span><span class="sxs-lookup"><span data-stu-id="281d6-152">If your add-in runs in one Office host, set the `Hosts` element in the manifest.</span></span> <span data-ttu-id="281d6-153">詳しくは、「[Hosts 要素を設定する](#set-the-hosts-element)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="281d6-153">For more information, see [Set the Hosts element](#set-the-hosts-element).</span></span>

- <span data-ttu-id="281d6-154">Office ホストがアドインを実行するためにサポートする必要のある最小要件セットまたは API メンバーを設定する`Requirements`には、マニフェストで要素を設定します。</span><span class="sxs-lookup"><span data-stu-id="281d6-154">To set the minimum requirement set or API members that an Office host must support to run your add-in, set the `Requirements` element in the manifest.</span></span> <span data-ttu-id="281d6-155">詳しくは、「[マニフェストで Requirements 要素を設定する](#set-the-requirements-element-in-the-manifest)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="281d6-155">For more information, see [Set the Requirements element in the manifest](#set-the-requirements-element-in-the-manifest).</span></span>

- <span data-ttu-id="281d6-156">Office ホストで特定の要件セットまたは API メンバーが利用可能である場合に追加の機能を提供する場合は、アドインの JavaScript コードでランタイム チェックを実行します。</span><span class="sxs-lookup"><span data-stu-id="281d6-156">If you would like to provide additional functionality if specific requirement sets or API members are available in the Office host, perform a runtime check in your add-in's JavaScript code.</span></span> <span data-ttu-id="281d6-157">たとえば、アドインが Excel 2016 で機能する場合は、Excel JavaScript API の API メンバーを使用して追加の機能を提供します。</span><span class="sxs-lookup"><span data-stu-id="281d6-157">For example, if your add-in runs in Excel 2016, use API members from the Excel JavaScript API to provide additional functionality.</span></span> <span data-ttu-id="281d6-158">詳細については、「[JavaScript コードでランタイム チェックを使用する](#use-runtime-checks-in-your-javascript-code)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="281d6-158">For more information, see [Use runtime checks in your JavaScript code](#use-runtime-checks-in-your-javascript-code).</span></span>

## <a name="set-the-hosts-element"></a><span data-ttu-id="281d6-159">Hosts 要素を設定する</span><span class="sxs-lookup"><span data-stu-id="281d6-159">Set the Hosts element</span></span>

<span data-ttu-id="281d6-160">1つの Office ホストアプリケーションでアドインを実行するには、マニフェスト`Hosts`内`Host`の要素と要素を使用します。</span><span class="sxs-lookup"><span data-stu-id="281d6-160">To make your add-in run in one Office host application, use the `Hosts` and `Host` elements in the manifest.</span></span> <span data-ttu-id="281d6-161">`Hosts`要素を指定しない場合、アドインはすべてのホストで実行されます。</span><span class="sxs-lookup"><span data-stu-id="281d6-161">If you don't specify the `Hosts` element, your add-in will run in all hosts.</span></span>

<span data-ttu-id="281d6-162">たとえば、次`Hosts`のと`Host`宣言は、アドインが excel のすべてのリリースで動作することを指定します。これには、Web、Windows、iPad 上の excel が含まれます。</span><span class="sxs-lookup"><span data-stu-id="281d6-162">For example, the following `Hosts` and `Host` declaration specifies that the add-in will work with any release of Excel, which includes Excel on the web, Windows, and iPad.</span></span>

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

<span data-ttu-id="281d6-163">要素`Hosts`には、1つ以上`Host`の要素を含めることができます。</span><span class="sxs-lookup"><span data-stu-id="281d6-163">The `Hosts` element can contain one or more `Host` elements.</span></span> <span data-ttu-id="281d6-164">要素`Host`は、アドインが必要とする Office ホストを指定します。</span><span class="sxs-lookup"><span data-stu-id="281d6-164">The `Host` element specifies the Office host your add-in requires.</span></span> <span data-ttu-id="281d6-165">`Name`属性は必須で、次のいずれかの値に設定できます。</span><span class="sxs-lookup"><span data-stu-id="281d6-165">The `Name` attribute is required and can be set to one of the following values.</span></span>

| <span data-ttu-id="281d6-166">名前</span><span class="sxs-lookup"><span data-stu-id="281d6-166">Name</span></span>          | <span data-ttu-id="281d6-167">Office ホスト アプリケーション</span><span class="sxs-lookup"><span data-stu-id="281d6-167">Office host applications</span></span>                                                                  |
|:--------------|:------------------------------------------------------------------------------------------|
| <span data-ttu-id="281d6-168">データベース</span><span class="sxs-lookup"><span data-stu-id="281d6-168">Database</span></span>      | <span data-ttu-id="281d6-169">Access Web アプリ</span><span class="sxs-lookup"><span data-stu-id="281d6-169">Access web apps</span></span>                                                                           |
| <span data-ttu-id="281d6-170">ドキュメント</span><span class="sxs-lookup"><span data-stu-id="281d6-170">Document</span></span>      | <span data-ttu-id="281d6-171">Windows 用 Word、Mac 用 Word、iPad 用 Word、Word on the web</span><span class="sxs-lookup"><span data-stu-id="281d6-171">Word on Windows, Word on Mac, Word on iPad, Word on the web</span></span>                               |
| <span data-ttu-id="281d6-172">Mailbox</span><span class="sxs-lookup"><span data-stu-id="281d6-172">Mailbox</span></span>       | <span data-ttu-id="281d6-173">Outlook on Windows、Outlook on Mac、Outlook on the web、Outlook on Android、Outlook on iOS</span><span class="sxs-lookup"><span data-stu-id="281d6-173">Outlook on Windows, Outlook on Mac, Outlook on the web, Outlook on Android, Outlook on iOS</span></span>|
| <span data-ttu-id="281d6-174">Presentation</span><span class="sxs-lookup"><span data-stu-id="281d6-174">Presentation</span></span>  | <span data-ttu-id="281d6-175">Windows 用 PowerPoint、Mac 用 PowerPoint、iPad 用 PowerPoint、PowerPoint on the web</span><span class="sxs-lookup"><span data-stu-id="281d6-175">PowerPoint on Windows, PowerPoint on Mac, PowerPoint on iPad, PowerPoint on the web</span></span>       |
| <span data-ttu-id="281d6-176">Project</span><span class="sxs-lookup"><span data-stu-id="281d6-176">Project</span></span>       | <span data-ttu-id="281d6-177">Windows での Project</span><span class="sxs-lookup"><span data-stu-id="281d6-177">Project on Windows</span></span>                                                                        |
| <span data-ttu-id="281d6-178">Workbook</span><span class="sxs-lookup"><span data-stu-id="281d6-178">Workbook</span></span>      | <span data-ttu-id="281d6-179">Windows 用 Excel、Mac 用 Excel、iPad 用 Excel、Excel on the web</span><span class="sxs-lookup"><span data-stu-id="281d6-179">Excel on Windows, Excel on Mac, Excel on iPad, Excel on the web</span></span>                           |

> [!NOTE]
> <span data-ttu-id="281d6-180">この`Name`属性は、アドインを実行できる Office ホストアプリケーションを指定します。</span><span class="sxs-lookup"><span data-stu-id="281d6-180">The `Name` attribute specifies the Office host application that can run your add-in.</span></span> <span data-ttu-id="281d6-181">Office ホストはさまざまなプラットフォームに対応しており、デスクトップ、Web ブラウザー、タブレット、モバイル デバイスで実行できます。</span><span class="sxs-lookup"><span data-stu-id="281d6-181">Office hosts are supported on different platforms and run on desktops, web browsers, tablets, and mobile devices.</span></span> <span data-ttu-id="281d6-182">アドインを実行するために使用するプラットフォームを指定することはできません。</span><span class="sxs-lookup"><span data-stu-id="281d6-182">You can't specify which platform can be used to run your add-in.</span></span> <span data-ttu-id="281d6-183">たとえば、`Mailbox` を指定した場合は、アドインの実行に Windows 用 Outlook と Outlook on the web の両方を使用できます。</span><span class="sxs-lookup"><span data-stu-id="281d6-183">For example, if you specify `Mailbox`, both Outlook on Windows and on the web can be used to run your add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="281d6-184">SharePoint で Access Web アプリとデータベースを作成して使用することは推奨されなくなりました。</span><span class="sxs-lookup"><span data-stu-id="281d6-184">We no longer recommend that you create and use Access web apps and databases in SharePoint.</span></span> <span data-ttu-id="281d6-185">代わりに、[Microsoft PowerApps](https://powerapps.microsoft.com/) を使用して、コード作成が不要な Web とモバイル デバイス用ビジネス ソリューションをビルドすることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="281d6-185">As an alternative, we recommend that you use [Microsoft PowerApps](https://powerapps.microsoft.com/) to build no-code business solutions for web and mobile devices.</span></span>


## <a name="set-the-requirements-element-in-the-manifest"></a><span data-ttu-id="281d6-186">マニフェストで Requirements 要素を設定する</span><span class="sxs-lookup"><span data-stu-id="281d6-186">Set the Requirements element in the manifest</span></span>

<span data-ttu-id="281d6-187">要素`Requirements`は、アドインを実行するために Office ホストでサポートされている必要のある最小要件セットまたは API メンバーを指定します。</span><span class="sxs-lookup"><span data-stu-id="281d6-187">The `Requirements` element specifies the minimum requirement sets or API members that must be supported by the Office host to run your add-in.</span></span> <span data-ttu-id="281d6-188">要素`Requirements`は、要件セットと、アドインで使用される個々のメソッドの両方を指定できます。</span><span class="sxs-lookup"><span data-stu-id="281d6-188">The `Requirements` element can specify both requirement sets and individual methods used in your add-in.</span></span> <span data-ttu-id="281d6-189">アドインマニフェストスキーマのバージョン1.1 では、Outlook アドインを`Requirements`除くすべてのアドインで、この要素は省略可能です。</span><span class="sxs-lookup"><span data-stu-id="281d6-189">In version 1.1 of the add-in manifest schema, the `Requirements` element is optional for all add-ins, except for Outlook add-ins.</span></span>

> [!WARNING]
> <span data-ttu-id="281d6-190">`Requirements`要素のみを使用して、アドインで使用する必要がある重要な要件セットまたは API メンバーを指定します。</span><span class="sxs-lookup"><span data-stu-id="281d6-190">Only use the `Requirements` element to specify critical requirement sets or API members that your add-in must use.</span></span> <span data-ttu-id="281d6-191">Office ホストまたはプラットフォームが、 `Requirements`要素で指定されている要件セットや API メンバーをサポートしていない場合、アドインはそのホストまたはプラットフォームでは実行されず、**アドイン**には表示されません。代わりに、web、Windows、iPad の Excel など、Office ホストのすべてのプラットフォームでアドインを使用できるようにすることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="281d6-191">If the Office host or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that host or platform, and won't display in **My Add-ins**. Instead, we recommend that you make your add-in available on all platforms of an Office host, such as Excel on the web, Windows, and iPad.</span></span> <span data-ttu-id="281d6-192">_すべて_の Office ホストおよびプラットフォームでアドインを使用できるようにするには、 `Requirements`要素の代わりにランタイムチェックを使用します。</span><span class="sxs-lookup"><span data-stu-id="281d6-192">To make your add-in available on  _all_ Office hosts and platforms, use runtime checks instead of the `Requirements` element.</span></span>

<span data-ttu-id="281d6-193">次のものをサポートするすべての Office ホスト アプリケーションで読み込まれるアドインのコード例を以下に示します。</span><span class="sxs-lookup"><span data-stu-id="281d6-193">The following code example shows an add-in that loads in all Office host applications that support the following:</span></span>

-  <span data-ttu-id="281d6-194">`TableBindings`要件セット。最小バージョンは "1.1" です。</span><span class="sxs-lookup"><span data-stu-id="281d6-194">`TableBindings` requirement set, which has a minimum version of "1.1".</span></span>

-  <span data-ttu-id="281d6-195">`OOXML`要件セット。最小バージョンは "1.1" です。</span><span class="sxs-lookup"><span data-stu-id="281d6-195">`OOXML` requirement set, which has a minimum version of "1.1".</span></span>

-  <span data-ttu-id="281d6-196">`Document.getSelectedDataAsync`手段.</span><span class="sxs-lookup"><span data-stu-id="281d6-196">`Document.getSelectedDataAsync` method.</span></span>

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

- <span data-ttu-id="281d6-197">要素`Requirements`には`Sets`および`Methods`子要素が含まれています。</span><span class="sxs-lookup"><span data-stu-id="281d6-197">The `Requirements` element contains the `Sets` and `Methods` child elements.</span></span>

- <span data-ttu-id="281d6-198">要素`Sets`には、1つ以上`Set`の要素を含めることができます。</span><span class="sxs-lookup"><span data-stu-id="281d6-198">The `Sets` element can contain one or more `Set` elements.</span></span> <span data-ttu-id="281d6-199">`DefaultMinVersion`すべての子`MinVersion` `Set`要素の既定値を指定します。</span><span class="sxs-lookup"><span data-stu-id="281d6-199">`DefaultMinVersion` specifies the default `MinVersion` value of all child `Set` elements.</span></span>

- <span data-ttu-id="281d6-200">要素`Set`は、Office ホストがアドインを実行するためにサポートする必要がある要件セットを指定します。</span><span class="sxs-lookup"><span data-stu-id="281d6-200">The `Set` element specifies requirement sets that the Office host must support to run the add-in.</span></span> <span data-ttu-id="281d6-201">属性`Name`は、要件セットの名前を指定します。</span><span class="sxs-lookup"><span data-stu-id="281d6-201">The `Name` attribute specifies the name of the requirement set.</span></span> <span data-ttu-id="281d6-202">は`MinVersion` 、要件セットの最小バージョンを指定します。</span><span class="sxs-lookup"><span data-stu-id="281d6-202">The `MinVersion` specifies the minimum version of the requirement set.</span></span> <span data-ttu-id="281d6-203">`MinVersion`API メンバーが属する`DefaultMinVersion`要件セットと要件セットのバージョンの詳細については、「 [Office アドインの要件セット](../reference/requirement-sets/office-add-in-requirement-sets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="281d6-203">`MinVersion` overrides the value of `DefaultMinVersion` For more information about requirement sets and requirement set versions that your API members belong to, see [Office Add-in requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md).</span></span>

- <span data-ttu-id="281d6-p122">要素`Methods`には、1つ以上`Method`の要素を含めることができます。Outlook アドインで`Methods`要素を使用することはできません。</span><span class="sxs-lookup"><span data-stu-id="281d6-p122">The `Methods` element can contain one or more `Method` elements. You can't use the `Methods` element with Outlook add-ins.</span></span>

- <span data-ttu-id="281d6-p123">要素`Method`は、アドインを実行する Office ホストでサポートする必要がある個別のメソッドを指定します。この`Name`属性は必須で、その親オブジェクトで修飾されたメソッドの名前を指定します。</span><span class="sxs-lookup"><span data-stu-id="281d6-p123">The `Method` element specifies an individual method that must be supported in the Office host where your add-in runs. The `Name` attribute is required and specifies the name of the method qualified with its parent object.</span></span>

## <a name="use-runtime-checks-in-your-javascript-code"></a><span data-ttu-id="281d6-208">JavaScript コードでランタイム チェックを使用する</span><span class="sxs-lookup"><span data-stu-id="281d6-208">Use runtime checks in your JavaScript code</span></span>

<span data-ttu-id="281d6-209">特定の要件セットが Office ホストでサポートされる場合、追加の機能を提供すると効果的な場合があります。</span><span class="sxs-lookup"><span data-stu-id="281d6-209">You might want to provide additional functionality in your add-in if certain requirement sets are supported by the Office host.</span></span> <span data-ttu-id="281d6-210">たとえば、アドインで Word 2016 を実行する場合、既存のアドインで Word JavaScript API を使用することがあります。</span><span class="sxs-lookup"><span data-stu-id="281d6-210">For example, you might want to use the Word JavaScript APIs in your existing add-in if your add-in runs in Word 2016.</span></span> <span data-ttu-id="281d6-211">その場合、要件セットの名前を指定し、[isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="281d6-211">To do this, you use the [isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) method with the name of the requirement set.</span></span> <span data-ttu-id="281d6-212">`isSetSupported`実行時に、アドインを実行している Office ホストが要件セットをサポートするかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="281d6-212">`isSetSupported` determines, at runtime, whether the Office host running the add-in supports the requirement set.</span></span> <span data-ttu-id="281d6-213">要件セットがサポートされて`isSetSupported`いる場合は、 **true**を返し、その要件セットの API メンバーを使用する追加のコードを実行します。</span><span class="sxs-lookup"><span data-stu-id="281d6-213">If the requirement set is supported, `isSetSupported` returns **true** and runs the additional code that uses the API members from that requirement set.</span></span> <span data-ttu-id="281d6-214">Office ホストが要件セットをサポートしてい`isSetSupported`ない場合は、 **false**を返し、追加のコードは実行されません。</span><span class="sxs-lookup"><span data-stu-id="281d6-214">If the Office host doesn't support the requirement set, `isSetSupported` returns **false** and the additional code won't run.</span></span> <span data-ttu-id="281d6-215">次のコードは、で`isSetSupported`使用する構文を示しています。</span><span class="sxs-lookup"><span data-stu-id="281d6-215">The following code shows the syntax to use with `isSetSupported`.</span></span>

```js
if (Office.context.requirements.isSetSupported(RequirementSetName, MinimumVersion))
{
   // Code that uses API members from RequirementSetName.
}

```

- <span data-ttu-id="281d6-216">_RequirementSetName_ (必須) は、要件セットの名前を表す文字列です (例: "**ExcelApi**"、"**Mailbox**" など)。</span><span class="sxs-lookup"><span data-stu-id="281d6-216">_RequirementSetName_ (required) is a string that represents the name of the requirement set (e.g., "**ExcelApi**", "**Mailbox**", etc.).</span></span> <span data-ttu-id="281d6-217">利用できる要件セットの詳細については、「[Office アドインの要件セット](../reference/requirement-sets/office-add-in-requirement-sets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="281d6-217">For more information about available requirement sets, see [Office Add-in requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md).</span></span>
- <span data-ttu-id="281d6-218">_MinimumVersion_ (省略可能) では、`if` ステートメントの範囲内でコードを実行するために、ホストがサポートする必要がある最小要件セットのバージョンを指定します (例: "**1.9**")。</span><span class="sxs-lookup"><span data-stu-id="281d6-218">_MinimumVersion_ (optional) is a string that specifies the minimum requirement set version that the host must support in order for the code within the `if` statement to run (e.g., "**1.9**").</span></span>

> [!WARNING]
> <span data-ttu-id="281d6-219">`isSetSupported`メソッドを呼び出す場合、 `MinimumVersion`パラメーター (指定されている場合) の値は文字列である必要があります。</span><span class="sxs-lookup"><span data-stu-id="281d6-219">When calling the `isSetSupported` method, the value of the `MinimumVersion` parameter (if specified) should be a string.</span></span> <span data-ttu-id="281d6-220">これは、JavaScript パーサーでは、1.1 や 1.10 のような数値の間の差異を区別できないが、"1.1" や "1.10" などの文字列値ではできるからです。</span><span class="sxs-lookup"><span data-stu-id="281d6-220">This is because the JavaScript parser cannot differentiate between numeric values such as 1.1 and 1.10, where as it can for string values such as "1.1" and "1.10".</span></span>
> <span data-ttu-id="281d6-221">`number` のオーバーロードは非推奨になります。</span><span class="sxs-lookup"><span data-stu-id="281d6-221">The `number` overload is deprecated.</span></span>

<span data-ttu-id="281d6-222">Office `isSetSupported`ホストと`RequirementSetName`関連付けられているを次のように使用します。</span><span class="sxs-lookup"><span data-stu-id="281d6-222">Use `isSetSupported` with the `RequirementSetName` associated with the Office host as follows.</span></span>

|<span data-ttu-id="281d6-223">Office ホスト</span><span class="sxs-lookup"><span data-stu-id="281d6-223">Office host</span></span>|<span data-ttu-id="281d6-224">RequirementSetName</span><span class="sxs-lookup"><span data-stu-id="281d6-224">RequirementSetName</span></span>|
|---|---|
|<span data-ttu-id="281d6-225">Excel</span><span class="sxs-lookup"><span data-stu-id="281d6-225">Excel</span></span>|<span data-ttu-id="281d6-226">ExcelApi</span><span class="sxs-lookup"><span data-stu-id="281d6-226">ExcelApi</span></span>|
|<span data-ttu-id="281d6-227">OneNote</span><span class="sxs-lookup"><span data-stu-id="281d6-227">OneNote</span></span>|<span data-ttu-id="281d6-228">OneNoteApi</span><span class="sxs-lookup"><span data-stu-id="281d6-228">OneNoteApi</span></span>|
|<span data-ttu-id="281d6-229">Outlook</span><span class="sxs-lookup"><span data-stu-id="281d6-229">Outlook</span></span>|<span data-ttu-id="281d6-230">Mailbox</span><span class="sxs-lookup"><span data-stu-id="281d6-230">Mailbox</span></span>|
|<span data-ttu-id="281d6-231">Word</span><span class="sxs-lookup"><span data-stu-id="281d6-231">Word</span></span>|<span data-ttu-id="281d6-232">WordApi</span><span class="sxs-lookup"><span data-stu-id="281d6-232">WordApi</span></span>|

<span data-ttu-id="281d6-233">これら`isSetSupported`のホストのメソッドと要件セットは、CDN の最新の Office .js ファイルで使用できます。</span><span class="sxs-lookup"><span data-stu-id="281d6-233">The `isSetSupported` method and the requirement sets for these hosts are available in the latest Office.js file on the CDN.</span></span> <span data-ttu-id="281d6-234">CDN から Office .js を使用しない場合は、例外が発生する可能性があるため`isSetSupported` 、アドインが例外を生成することがあります。</span><span class="sxs-lookup"><span data-stu-id="281d6-234">If you don't use Office.js from the CDN, your add-in might generate exceptions because `isSetSupported` will be undefined.</span></span> <span data-ttu-id="281d6-235">詳細については、「[最新の Office JAVASCRIPT API ライブラリを指定する](#specify-the-latest-office-javascript-api-library)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="281d6-235">For more information, see [Specify the latest Office JavaScript API library](#specify-the-latest-office-javascript-api-library).</span></span>

<span data-ttu-id="281d6-236">次のコードの例は、さまざまな要件セットや API メンバーをサポートするさまざまな Office ホストにおいて、アドインで各種の機能を提供する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="281d6-236">The following code example shows how an add-in can provide different functionality for different Office hosts that might support different requirement sets or API members.</span></span>

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
    // Run additional code when the Office host is not Word 2016 or later and does not support the CustomXmlParts requirement set.
}

```

## <a name="runtime-checks-using-methods-not-in-a-requirement-set"></a><span data-ttu-id="281d6-237">要件セットにないメソッドを使用したランタイム チェック</span><span class="sxs-lookup"><span data-stu-id="281d6-237">Runtime checks using methods not in a requirement set</span></span>

<span data-ttu-id="281d6-238">API の一部のメンバーは、要件のセットに属していません。</span><span class="sxs-lookup"><span data-stu-id="281d6-238">Some API members don't belong to requirement sets.</span></span> <span data-ttu-id="281d6-239">これは、 [Office JavaScript api](../reference/javascript-api-for-office.md)名前空間 ( `Office.` [Outlook メールボックス api](/javascript/api/outlook)以外のすべて) に属する api メンバーではなく、 [Word javascript api](../reference/overview/word-add-ins-reference-overview.md) (すべて`Word.`のもの)、 [Excel javascript api](../reference/overview/excel-add-ins-reference-overview.md) `Excel.`(すべての場合)、または[OneNote javascript api](../reference/overview/onenote-add-ins-javascript-reference.md) ( `OneNote.`あらゆる場合) の名前空間に含まれる api メンバーにのみ適用されます。</span><span class="sxs-lookup"><span data-stu-id="281d6-239">This only applies to API members that are part of the [Office JavaScript API](../reference/javascript-api-for-office.md) namespace (anything under `Office.` except [Outlook Mailbox APIs](/javascript/api/outlook)), but not API members that belong to the [Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) (anything in `Word.`), [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) (anything in `Excel.`), or [OneNote JavaScript API](../reference/overview/onenote-add-ins-javascript-reference.md) (anything in `OneNote.`) namespaces.</span></span> <span data-ttu-id="281d6-240">要件セットに属さないメソッドにアドインが依存するとき、ランタイム チェックを利用し、メソッドが Office ホストでサポートされているかどうかを判断できます。たとえば、次のコード例のようになります。</span><span class="sxs-lookup"><span data-stu-id="281d6-240">When your add-in depends on a method that is not part of a requirement set, you can use the runtime check to determine whether the method is supported by the Office host, as shown in the following code example.</span></span> <span data-ttu-id="281d6-241">要件セットに属さないメソッドの詳細な一覧については、「[Office アドインの要件セット](../reference/requirement-sets/office-add-in-requirement-sets.md#methods-that-arent-part-of-a-requirement-set)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="281d6-241">For a complete list of methods that don't belong to a requirement set, see [Office Add-in requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md#methods-that-arent-part-of-a-requirement-set).</span></span>

> [!NOTE]
> <span data-ttu-id="281d6-242">アドインのコードでのこの種のランタイム チェックは、限定的に使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="281d6-242">We recommend that you limit the use of this type of runtime check in your add-in's code.</span></span>

<span data-ttu-id="281d6-243">次のコード例では、ホストが`document.setSelectedDataAsync`サポートしているかどうかを確認します。</span><span class="sxs-lookup"><span data-stu-id="281d6-243">The following code example checks whether the host supports `document.setSelectedDataAsync`.</span></span>

```js
if (Office.context.document.setSelectedDataAsync)
{
    // Run code that uses document.setSelectedDataAsync.
}
```


## <a name="see-also"></a><span data-ttu-id="281d6-244">関連項目</span><span class="sxs-lookup"><span data-stu-id="281d6-244">See also</span></span>

- [<span data-ttu-id="281d6-245">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="281d6-245">Office Add-ins XML manifest</span></span>](add-in-manifests.md)
- [<span data-ttu-id="281d6-246">Office アドインの要件セット</span><span class="sxs-lookup"><span data-stu-id="281d6-246">Office Add-in requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="281d6-247">Word-Add-in-Get-Set-EditOpen-XML</span><span class="sxs-lookup"><span data-stu-id="281d6-247">Word-Add-in-Get-Set-EditOpen-XML</span></span>](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML)
