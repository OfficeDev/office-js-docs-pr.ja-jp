---
title: Office のホストと API の要件を指定する
description: ''
ms.date: 05/29/2019
localization_priority: Priority
ms.openlocfilehash: ccff7ba1896c9d1683f9fc9d67cdd79fe52da623
ms.sourcegitcommit: b299b8a5dfffb6102cb14b431bdde4861abfb47f
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/30/2019
ms.locfileid: "34589147"
---
# <a name="specify-office-hosts-and-api-requirements"></a><span data-ttu-id="e112c-102">Office のホストと API の要件を指定する</span><span class="sxs-lookup"><span data-stu-id="e112c-102">Specify Office hosts and API requirements</span></span>

<span data-ttu-id="e112c-p101">期待どおりの動作をするうえで、Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。たとえば、次のようなアドインがあります。</span><span class="sxs-lookup"><span data-stu-id="e112c-p101">Your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API in order to work as expected. For example, your add-in might:</span></span>

- <span data-ttu-id="e112c-105">1 つの Office アプリケーション (例: Word または Excel)、またはいくつかのアプリケーションで実行します。</span><span class="sxs-lookup"><span data-stu-id="e112c-105">Run in a single Office application (Word or Excel), or several applications.</span></span>

- <span data-ttu-id="e112c-p102">Office の一部のバージョンでのみ利用できる JavaScript API を使用します。たとえば、Excel 2016 で実行するアドインでは、Excel JavaScript API を使用することがあります。</span><span class="sxs-lookup"><span data-stu-id="e112c-p102">Make use of JavaScript APIs that are only available in some versions of Office. For example, you might use the Excel JavaScript APIs in an add-in that runs in Excel 2016.</span></span>

- <span data-ttu-id="e112c-108">アドインが使用する API メンバーをサポートするバージョンの Office でのみ実行します。</span><span class="sxs-lookup"><span data-stu-id="e112c-108">Run only in versions of Office that support API members that your add-in uses.</span></span>

<span data-ttu-id="e112c-109">この記事は、アドインが期待どおりに機能し、できるだけ多くのユーザーが利用できるようにするために選択する必要のあるオプションについて理解するのに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="e112c-109">This article helps you understand which options you should choose to ensure that your add-in works as expected and reaches the broadest audience possible.</span></span>

> [!NOTE]
> <span data-ttu-id="e112c-110">現時点での Office アドインのサポート状況の概要については、「[Office アドインを使用できるホストおよびプラットフォーム](../overview/office-add-in-availability.md)」のページを参照してください。</span><span class="sxs-lookup"><span data-stu-id="e112c-110">For a high-level view of where Office Add-ins are currently supported, see the [Office Add-in host and platform availability](../overview/office-add-in-availability.md) page.</span></span>

<span data-ttu-id="e112c-111">この記事で説明する中心的な概念を次の表に示します。</span><span class="sxs-lookup"><span data-stu-id="e112c-111">The following table lists core concepts discussed throughout this article.</span></span>

|<span data-ttu-id="e112c-112">**概念**</span><span class="sxs-lookup"><span data-stu-id="e112c-112">**Concept**</span></span>|<span data-ttu-id="e112c-113">**説明**</span><span class="sxs-lookup"><span data-stu-id="e112c-113">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="e112c-114">Office アプリケーション、Office ホスト アプリケーション、Office ホスト、またはホスト</span><span class="sxs-lookup"><span data-stu-id="e112c-114">Office application, Office host application, Office host, or host</span></span>|<span data-ttu-id="e112c-p103">アドインを実行するために使用される Office アプリケーション。たとえば、Word、Word Online、Excel など。</span><span class="sxs-lookup"><span data-stu-id="e112c-p103">The Office application used to run your add-in. For example, Word, Word Online, Excel, and so on.</span></span>|
|<span data-ttu-id="e112c-117">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="e112c-117">Platform</span></span>|<span data-ttu-id="e112c-118">Office Online、Office for iPad などの Office ホストを実行する場所。</span><span class="sxs-lookup"><span data-stu-id="e112c-118">Where the Office host runs, such as Office Online or Office for iPad.</span></span>|
|<span data-ttu-id="e112c-119">要件セット</span><span class="sxs-lookup"><span data-stu-id="e112c-119">Requirement set</span></span>|<span data-ttu-id="e112c-p104">関連する API メンバーの名前付きグループ。アドインは要件セットを使用して、Office ホストが、アドインによって使用される API メンバーをサポートしているかどうかを判別します。個々の API メンバーのサポートをテストするよりも、要件セットのサポートをテストするほうが簡単です。要件セットのサポートは、Office ホストと Office ホストのバージョンによって異なります。 </span><span class="sxs-lookup"><span data-stu-id="e112c-p104">A named group of related API members. Add-ins use requirement sets to determine whether the Office host supports API members used by your add-in. It's easier to test for the support of a requirement set than for the support of individual API members. Requirement set support varies by Office host and the version of the Office host. </span></span><br ><span data-ttu-id="e112c-124">要件セットはマニフェスト ファイルで指定されます。</span><span class="sxs-lookup"><span data-stu-id="e112c-124">Requirement sets are specified in the manifest file.</span></span> <span data-ttu-id="e112c-125">マニフェストで要件セットを指定するときは、アドインを実行するために Office ホストが提供する必要のある最小レベルの API サポートを設定します。</span><span class="sxs-lookup"><span data-stu-id="e112c-125">When you specify requirement sets in the manifest, you set the minimum level of API support that the Office host must provide in order to run your add-in.</span></span> <span data-ttu-id="e112c-126">マニフェストで指定されている要件セットをサポートしていない Office ホストはアドインを実行できず、アドインは <span class="ui">[個人用アドイン]</span> に表示されません。これにより、アドインが利用できる場所が制限されます。</span><span class="sxs-lookup"><span data-stu-id="e112c-126">Office hosts that don't support requirement sets specified in the manifest can't run your add-in, and your add-in won't display in <span class="ui">My Add-ins</span>. This restricts where your add-in is available.In code using runtime checks.</span></span> <span data-ttu-id="e112c-127">コードでは、ランタイム チェックを使用します。</span><span class="sxs-lookup"><span data-stu-id="e112c-127">In code using runtime checks.</span></span> <span data-ttu-id="e112c-128">要件セットの詳細な一覧については、「[Office アドインの要件セット](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e112c-128">For the complete list of requirement sets, see [Office Add-in requirement sets](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span></span>|
|<span data-ttu-id="e112c-129">ランタイム チェック</span><span class="sxs-lookup"><span data-stu-id="e112c-129">Runtime check</span></span>|<span data-ttu-id="e112c-p106">アドインを実行している Office ホストが、アドインで使用されている要件セットまたはメソッドをサポートしているかどうかを判別するために実行時に行われるテスト。ランタイム チェックを実行するには、**if** ステートメントに **isSetSupported** メソッド、要件セット、または要件セットの一部ではないメソッド名を指定して使用します。ランタイム チェックを使用し、多くのユーザーが対象のアドインを使用できることを確認します。要件セットとは異なり、ランタイム チェックでは、対象アドインを実行するために Office ホストが提供する必要のある最小レベルの API サポートは指定しません。代わりに、**if** ステートメントを使用して API メンバーがサポートされているかどうかを判別します。サポートされている場合には、アドインで追加機能を提供できます。ランタイム チェックを使用するときは、自分のアドインは必ず **[個人用アドイン]** に表示されます。</span><span class="sxs-lookup"><span data-stu-id="e112c-p106">A test that is performed at runtime to determine whether the Office host running your add-in supports requirement sets or methods used by your add-in. To perform a runtime check, you use an  **if** statement with the **isSetSupported** method, the requirement sets, or the method names that aren't part of a requirement set.Use runtime checks to ensure that your add-in reaches the broadest number of customers. Unlike requirement sets, runtime checks don't specify the minimum level of API support that the Office host must provide for your add-in to run. Instead, you use the  **if** statement to determine whether an API member is supported. If it is, you can provide additional functionality in your add-in. Your add-in will always display in **My Add-ins** when you use runtime checks.</span></span>|

## <a name="before-you-begin"></a><span data-ttu-id="e112c-136">始める前に</span><span class="sxs-lookup"><span data-stu-id="e112c-136">Before you begin</span></span>

<span data-ttu-id="e112c-p107">アドインで最新バージョンのアドイン マニフェスト スキーマを使用する必要があります。アドインでランタイム チェックを使用する場合は、最新の JavaScript API for Office (office.js) ライブラリを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e112c-p107">Your add-in must use the most current version of the add-in manifest schema. If you use runtime checks in your add-in, ensure that you use the latest JavaScript API for Office (office.js) library.</span></span>

### <a name="specify-the-latest-add-in-manifest-schema"></a><span data-ttu-id="e112c-139">最新のアドイン マニフェスト スキーマを指定する</span><span class="sxs-lookup"><span data-stu-id="e112c-139">Specify the latest add-in manifest schema</span></span>

<span data-ttu-id="e112c-p108">アドインのマニフェストでは、アドイン マニフェスト スキーマのバージョン 1.1 を使用する必要があります。アドイン マニフェストの **OfficeApp** 要素を次のように設定します。</span><span class="sxs-lookup"><span data-stu-id="e112c-p108">Your add-in's manifest must use version 1.1 of the add-in manifest schema. Set the  **OfficeApp** element in your add-in manifest as follows.</span></span>

```XML
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
```

### <a name="specify-the-latest-javascript-api-for-office-library"></a><span data-ttu-id="e112c-142">最新の JavaScript API for Office ライブラリを指定する</span><span class="sxs-lookup"><span data-stu-id="e112c-142">Specify the latest JavaScript API for Office library</span></span>

<span data-ttu-id="e112c-p109">ランタイム チェックを使用する場合、コンテンツ配信ネットワーク (CDN) から JavaScript API for Office ライブラリの最新版を参照します。その場合、HTML に次の `script` タグを追加します。CDN URL で `/1/` を使用することで、Office.js の最新版が参照されます。</span><span class="sxs-lookup"><span data-stu-id="e112c-p109">If you use runtime checks, reference the most current version of the JavaScript API for Office library from the content delivery network (CDN). To do this, add the following  `script` tag to your HTML. Using `/1/` in the CDN URL ensures that you reference the most recent version of Office.js.</span></span>

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

## <a name="options-to-specify-office-hosts-or-api-requirements"></a><span data-ttu-id="e112c-146">Office のホストや API の要件を指定するオプション</span><span class="sxs-lookup"><span data-stu-id="e112c-146">Options to specify Office hosts or API requirements</span></span>

<span data-ttu-id="e112c-p110">Office ホストまたは API の要件を指定するときに、検討すべき事項がいくつかあります。次の図に、アドインで使用すべき手法の判別方法を示します。</span><span class="sxs-lookup"><span data-stu-id="e112c-p110">When you specify Office hosts or API requirements, there are several factors to consider. The following diagram shows how to decide which technique to use in your add-in.</span></span>

![Office のホストまたは API の要件を指定する際に、アドインに最適なオプションを選択する](../images/options-for-office-hosts.png)

- <span data-ttu-id="e112c-p111">アドインを 1 つの Office ホストで実行する場合、マニフェストに **Hosts** 要素を設定します。詳しくは、「[Hosts 要素を設定する](#set-the-hosts-element)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e112c-p111">If your add-in runs in one Office host, set the **Hosts** element in the manifest. For more information, see [Set the Hosts element](#set-the-hosts-element).</span></span>

- <span data-ttu-id="e112c-p112">アドインを実行するために Office ホストがサポートする必要のある最小レベルの要件セットまたは API メンバーを設定するには、マニフェストに **Requirements** 要素を設定します。詳しくは、「[マニフェストで Requirements 要素を設定する](#set-the-requirements-element-in-the-manifest)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="e112c-p112">To set the minimum requirement set or API members that an Office host must support to run your add-in, set the  **Requirements** element in the manifest. For more information, see [Set the Requirements element in the manifest](#set-the-requirements-element-in-the-manifest).</span></span>

- <span data-ttu-id="e112c-154">Office ホストで特定の要件セットまたは API メンバーが利用可能である場合に追加の機能を提供する場合は、アドインの JavaScript コードでランタイム チェックを実行します。</span><span class="sxs-lookup"><span data-stu-id="e112c-154">If you would like to provide additional functionality if specific requirement sets or API members are available in the Office host, perform a runtime check in your add-in's JavaScript code.</span></span> <span data-ttu-id="e112c-155">たとえば、アドインが Excel 2016 で機能する場合は、Excel JavaScript API の API メンバーを使用して追加の機能を提供します。</span><span class="sxs-lookup"><span data-stu-id="e112c-155">For example, if your add-in runs in Excel 2016, use API members from the new JavaScript API for Excel to provide additional functionality.</span></span> <span data-ttu-id="e112c-156">詳細については、「[JavaScript コードでランタイム チェックを使用する](#use-runtime-checks-in-your-javascript-code)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="e112c-156">For more information, see [Use runtime checks in your JavaScript code](#use-runtime-checks-in-your-javascript-code).</span></span>

## <a name="set-the-hosts-element"></a><span data-ttu-id="e112c-157">Hosts 要素を設定する</span><span class="sxs-lookup"><span data-stu-id="e112c-157">Set the Hosts element</span></span>

<span data-ttu-id="e112c-p114">アドインを 1 つの Office ホスト アプリケーションで実行するには、マニフェストで **Hosts** 要素と **Host** 要素を使用します。**Hosts** 要素を指定しない場合、アドインはすべてのホストで実行されます。</span><span class="sxs-lookup"><span data-stu-id="e112c-p114">To make your add-in run in one Office host application, use the  **Hosts** and **Host** elements in the manifest. If you don't specify the **Hosts** element, your add-in will run in all hosts.</span></span>

<span data-ttu-id="e112c-160">たとえば、次の **Hosts** と **Host** の宣言は、アドインが Excel のすべてのリリース (これには、Windows での Excel、Excel Online、Excel for iPad が含まれる) で機能することを指定しています。</span><span class="sxs-lookup"><span data-stu-id="e112c-160">For example, the following  **Hosts** and **Host** declaration specifies that the add-in will work with any release of Excel, which includes Excel on Windows, Excel Online, and Excel for iPad.</span></span>

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

<span data-ttu-id="e112c-p115">**Hosts** 要素には、1 つ以上の **Host** 要素を含めることができます。**Host** 要素は、アドインで必要な Office ホストを指定します。**Name** 属性は必須で、次のいずれかの値に設定できます。</span><span class="sxs-lookup"><span data-stu-id="e112c-p115">The  **Hosts** element can contain one or more **Host** elements. The **Host** element specifies the Office host your add-in requires. The **Name** attribute is required and can be set to one of the following values.</span></span>

| <span data-ttu-id="e112c-164">名前</span><span class="sxs-lookup"><span data-stu-id="e112c-164">Name</span></span>          | <span data-ttu-id="e112c-165">Office ホスト アプリケーション</span><span class="sxs-lookup"><span data-stu-id="e112c-165">Office host applications</span></span>                                                              |
|:--------------|:--------------------------------------------------------------------------------------|
| <span data-ttu-id="e112c-166">データベース</span><span class="sxs-lookup"><span data-stu-id="e112c-166">Database</span></span>      | <span data-ttu-id="e112c-167">Access Web アプリ</span><span class="sxs-lookup"><span data-stu-id="e112c-167">Access web apps</span></span>                                                                       |
| <span data-ttu-id="e112c-168">Document</span><span class="sxs-lookup"><span data-stu-id="e112c-168">Document</span></span>      | <span data-ttu-id="e112c-169">Windows での Word、Word for Mac、Word for iPad、Word Online</span><span class="sxs-lookup"><span data-stu-id="e112c-169">Word on Windows, word for Mac, Word for iPad, and Word Online</span></span>                         |
| <span data-ttu-id="e112c-170">メールボックス</span><span class="sxs-lookup"><span data-stu-id="e112c-170">Mailbox</span></span>       | <span data-ttu-id="e112c-171">Windows での Outlook、Outlook for Mac、Outlook on the web、Outlook.com</span><span class="sxs-lookup"><span data-stu-id="e112c-171">Outlook on Windows, Outlook for Mac, Outlook on the web, and Outlook.com</span></span>              |
| <span data-ttu-id="e112c-172">プレゼンテーション</span><span class="sxs-lookup"><span data-stu-id="e112c-172">Presentation</span></span>  | <span data-ttu-id="e112c-173">Windows での PowerPoint、PowerPoint for Mac、PowerPoint for iPad、PowerPoint Online</span><span class="sxs-lookup"><span data-stu-id="e112c-173">PowerPoint on Windows, PowerPoint for Mac, PowerPoint for iPad, and PowerPoint Online</span></span> |
| <span data-ttu-id="e112c-174">Project</span><span class="sxs-lookup"><span data-stu-id="e112c-174">Project</span></span>       | <span data-ttu-id="e112c-175">Windows での Project</span><span class="sxs-lookup"><span data-stu-id="e112c-175">Project 2016 or later on Windows</span></span>                                                                    |
| <span data-ttu-id="e112c-176">ブック</span><span class="sxs-lookup"><span data-stu-id="e112c-176">Workbook</span></span>      | <span data-ttu-id="e112c-177">Windows での Excel、Excel for Mac、Excel for iPad、Excel Online</span><span class="sxs-lookup"><span data-stu-id="e112c-177">Excel on Windows, Excel for Mac, Excel for iPad, and Excel Online</span></span>                     |

> [!NOTE]
> <span data-ttu-id="e112c-p116">`Name` 属性により、アドインを実行できる Office ホスト アプリケーションが指定されます。Office ホストはさまざまなプラットフォームに対応しており、デスクトップ、Web ブラウザー、タブレット、モバイル デバイスで実行できます。アドインを実行するために使用するプラットフォームを指定することはできません。たとえば、`Mailbox` を指定した場合、Outlook と Outlook Web App の両方を利用してアドインを実行できます。</span><span class="sxs-lookup"><span data-stu-id="e112c-p116">The  `Name` attribute specifies the Office host application that can run your add-in. Office hosts are supported on different platforms and run on desktops, web browsers, tablets, and mobile devices. You can't specify which platform can be used to run your add-in. For example, if you specify `Mailbox`, both Outlook and Outlook Web App can be used to run your add-in.</span></span>


## <a name="set-the-requirements-element-in-the-manifest"></a><span data-ttu-id="e112c-182">マニフェストで Requirements 要素を設定する</span><span class="sxs-lookup"><span data-stu-id="e112c-182">Set the Requirements element in the manifest</span></span>

<span data-ttu-id="e112c-p117">**Requirements** 要素は、アドインを実行するために Office ホストによってサポートされている必要がある最小要件セットまたは API メンバーを指定します。**Requirements** 要素は、アドインで使用される要件セットと個々のメソッドの両方を指定できます。アドイン マニフェスト スキーマのバージョン 1.1 では、**Requirements** 要素は Outlook アドイン以外のすべてのアドインで省略可能です。</span><span class="sxs-lookup"><span data-stu-id="e112c-p117">The  **Requirements** element specifies the minimum requirement sets or API members that must be supported by the Office host to run your add-in. The **Requirements** element can specify both requirement sets and individual methods used in your add-in. In version 1.1 of the add-in manifest schema, the **Requirements** element is optional for all add-ins, except for Outlook add-ins.</span></span>

> [!WARNING]
> <span data-ttu-id="e112c-p118">アドインで必須の重要な要件セットまたは API メンバーを指定するには、**Requirements** 要素のみを使用します。Office ホストまたはプラットフォームが、**Requirements** 要素で指定した要件セットまたは API メンバーをサポートしない場合、アドインはそのホストまたはプラットフォームでは実行されず、**[個人用アドイン]** にも表示されません。代わりに、Windows での Excel、Excel Online、Excel for iPad などの Office ホストのすべてのプラットフォームでアドインを使用できるようにすることをお勧めします。_すべて_の Office ホストとプラットフォームでアドインを使用できるようにするには、**Requirements** 要素ではなく、ランタイム チェックを使用します。</span><span class="sxs-lookup"><span data-stu-id="e112c-p118">Only use the **Requirements** element to specify critical requirement sets or API members that your add-in must use. If the Office host or platform doesn't support the requirement sets or API members specified in the **Requirements** element, the add-in won't run in that host or platform, and won't display in **My Add-ins**. Instead, we recommend that you make your add-in available on all platforms of an Office host, such as Excel on Windows, Excel Online, and Excel for iPad. To make your add-in available on  _all_ Office hosts and platforms, use runtime checks instead of the **Requirements** element.</span></span>

<span data-ttu-id="e112c-189">次のものをサポートするすべての Office ホスト アプリケーションで読み込まれるアドインのコード例を以下に示します。</span><span class="sxs-lookup"><span data-stu-id="e112c-189">The following code example shows an add-in that loads in all Office host applications that support the following:</span></span>

-  <span data-ttu-id="e112c-190">**TableBindings** 要件セット。最小バージョンは 1.1。</span><span class="sxs-lookup"><span data-stu-id="e112c-190">**TableBindings** requirement set, which has a minimum version of 1.1.</span></span>

-  <span data-ttu-id="e112c-191">**OOXML** 要件セット。最小バージョンは 1.1。</span><span class="sxs-lookup"><span data-stu-id="e112c-191">**OOXML** requirement set, which has a minimum version of 1.1.</span></span>

-  <span data-ttu-id="e112c-192">**Document.getSelectedDataAsync** メソッド。</span><span class="sxs-lookup"><span data-stu-id="e112c-192">**Document.getSelectedDataAsync** method.</span></span>

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

- <span data-ttu-id="e112c-193">**Requirements** 要素には **Sets** 子要素と **Methods** 子要素が含まれます。</span><span class="sxs-lookup"><span data-stu-id="e112c-193">The  **Requirements** element contains the **Sets** and **Methods** child elements.</span></span>

- <span data-ttu-id="e112c-p119">**Sets** 要素には、1 つ以上の **Set** 要素を含めることができます。**DefaultMinVersion** は、すべての **Set** 子要素の **MinVersion** の既定値を指定します。</span><span class="sxs-lookup"><span data-stu-id="e112c-p119">The  **Sets** element can contain one or more **Set** elements. **DefaultMinVersion** specifies the default **MinVersion** value of all child **Set** elements.</span></span>

- <span data-ttu-id="e112c-196">**Set** 要素は、アドインを実行するために Office ホストがサポートする必要のある要件セットを指定します。</span><span class="sxs-lookup"><span data-stu-id="e112c-196">The  **Set** element specifies requirement sets that the Office host must support to run the add-in.</span></span> <span data-ttu-id="e112c-197">**Name** 属性は、要件セットの名前を指定します。</span><span class="sxs-lookup"><span data-stu-id="e112c-197">The **Name** attribute specifies the name of the requirement set.</span></span> <span data-ttu-id="e112c-198">**MinVersion** は、要件セットの最小バージョンを指定します。</span><span class="sxs-lookup"><span data-stu-id="e112c-198">The **MinVersion** specifies the minimum version of the requirement set.</span></span> <span data-ttu-id="e112c-199">**MinVersion** は、**DefaultMinVersion** の値よりも優先されます。</span><span class="sxs-lookup"><span data-stu-id="e112c-199">**MinVersion** overrides the value of **DefaultMinVersion**.</span></span> <span data-ttu-id="e112c-200">要件セットと API メンバーが属する要件セットのバージョンの詳細については、「[「Office アドインの要件セット](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e112c-200">For more information about requirement sets and requirement set versions that your API members belong to, see [Office Add-in requirement sets](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span></span>

- <span data-ttu-id="e112c-p121">**Methods** 要素には、1 つ以上の **Method** 要素を含めることができます。Outlook アドインで **Methods** 要素を使用することはできません。</span><span class="sxs-lookup"><span data-stu-id="e112c-p121">The  **Methods** element can contain one or more **Method** elements. You can't use the **Methods** element with Outlook add-ins.</span></span>

- <span data-ttu-id="e112c-p122">**Method** 要素は、アドインが実行される Office ホストでサポートされている必要のある個々のメソッドを指定します。**Name** 属性は必須であり、親オブジェクトで修飾されたメソッドの名前を指定します。</span><span class="sxs-lookup"><span data-stu-id="e112c-p122">The  **Method** element specifies an individual method that must be supported in the Office host where your add-in runs. The **Name** attribute is required and specifies the name of the method qualified with its parent object.</span></span>


## <a name="use-runtime-checks-in-your-javascript-code"></a><span data-ttu-id="e112c-205">JavaScript コードでランタイム チェックを使用する</span><span class="sxs-lookup"><span data-stu-id="e112c-205">Use runtime checks in your JavaScript code</span></span>


<span data-ttu-id="e112c-206">特定の要件セットが Office ホストでサポートされる場合、追加の機能を提供すると効果的な場合があります。</span><span class="sxs-lookup"><span data-stu-id="e112c-206">You might want to provide additional functionality in your add-in if certain requirement sets are supported by the Office host.</span></span> <span data-ttu-id="e112c-207">たとえば、アドインで Word 2016 を実行する場合、既存のアドインで Word JavaScript API を使用することがあります。</span><span class="sxs-lookup"><span data-stu-id="e112c-207">For example, you might want to use the new Word JavaScript APIs Word in your existing add-in if your add-in runs in Word 2016.</span></span> <span data-ttu-id="e112c-208">その場合、要件セットの名前を指定し、[isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="e112c-208">To do this, you use the  [isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) method with the name of the requirement set.</span></span> <span data-ttu-id="e112c-209">**isSetSupported** により実行時に、アドインを実行する Office ホストが要件セットをサポートするかどうかが判断されます。</span><span class="sxs-lookup"><span data-stu-id="e112c-209">**isSetSupported** determines, at runtime, whether the Office host running the add-in supports the requirement set.</span></span> <span data-ttu-id="e112c-210">要件セットがサポートされる場合、**isSetSupported** は **true** を返し、その要件セットから API メンバーを使用する追加コードを実行します。</span><span class="sxs-lookup"><span data-stu-id="e112c-210">If the requirement set is supported, **isSetSupported** returns **true** and runs the additional code that uses the API members from that requirement set.</span></span> <span data-ttu-id="e112c-211">Office ホストで要件セットがサポートされない場合、**isSetSupported** は **false** を返し、追加コードは実行されません。</span><span class="sxs-lookup"><span data-stu-id="e112c-211">If the Office host doesn't support the requirement set, **isSetSupported** returns **false** and the additional code won't run.</span></span> <span data-ttu-id="e112c-212">次のコードは、**isSetSupported** と共に使用する構文を示しています。</span><span class="sxs-lookup"><span data-stu-id="e112c-212">The following code shows the syntax to use with **isSetSupported**.</span></span>


```js
if (Office.context.requirements.isSetSupported(RequirementSetName, VersionNumber))
{
   // Code that uses API members from RequirementSetName.
}

```

-  <span data-ttu-id="e112c-213">_RequirementSetName_ (必須) は、要件セットの名前を表す文字列です。</span><span class="sxs-lookup"><span data-stu-id="e112c-213">_RequirementSetName_ (required) is a string that represents the name of the requirement set.</span></span> <span data-ttu-id="e112c-214">利用できる要件セットの詳細については、「[Office アドインの要件セット](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e112c-214">For more information about available requirement sets, see [Office Add-in requirement sets](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span></span>
    
-  <span data-ttu-id="e112c-215">_VersionNumber_ (省略可能) は要件セットのバージョンです。</span><span class="sxs-lookup"><span data-stu-id="e112c-215">_VersionNumber_ (optional) is the version of the requirement set.</span></span>

<span data-ttu-id="e112c-216">次のように、Office ホストに関連付けられている **RequirementSetName** と一緒に **isSetSupported** を使用します。</span><span class="sxs-lookup"><span data-stu-id="e112c-216">Use **isSetSupported** with the **RequirementSetName** associated with the Office host as follows.</span></span>

|<span data-ttu-id="e112c-217">Office ホスト</span><span class="sxs-lookup"><span data-stu-id="e112c-217">Office host</span></span>|<span data-ttu-id="e112c-218">RequirementSetName</span><span class="sxs-lookup"><span data-stu-id="e112c-218">RequirementSetName</span></span>|
|---|---|
|<span data-ttu-id="e112c-219">Excel</span><span class="sxs-lookup"><span data-stu-id="e112c-219">Excel</span></span>|<span data-ttu-id="e112c-220">ExcelApi</span><span class="sxs-lookup"><span data-stu-id="e112c-220">ExcelApi</span></span>|
|<span data-ttu-id="e112c-221">OneNote</span><span class="sxs-lookup"><span data-stu-id="e112c-221">OneNote</span></span>|<span data-ttu-id="e112c-222">OneNoteApi</span><span class="sxs-lookup"><span data-stu-id="e112c-222">OneNoteApi</span></span>|
|<span data-ttu-id="e112c-223">Outlook</span><span class="sxs-lookup"><span data-stu-id="e112c-223">Outlook</span></span>|<span data-ttu-id="e112c-224">Mailbox</span><span class="sxs-lookup"><span data-stu-id="e112c-224">Mailbox</span></span>|
|<span data-ttu-id="e112c-225">Word</span><span class="sxs-lookup"><span data-stu-id="e112c-225">Word</span></span>|<span data-ttu-id="e112c-226">WordApi</span><span class="sxs-lookup"><span data-stu-id="e112c-226">WordApi</span></span>|

<span data-ttu-id="e112c-227">**isSetSupported** メソッドおよびこれらの要件セットは、CDN の最新の Office.js ファイルで利用できます。</span><span class="sxs-lookup"><span data-stu-id="e112c-227">The **isSetSupported** method, and the ExcelAPI and WordAPI requirement sets, are available in the latest Office.js file available from the CDN.</span></span> <span data-ttu-id="e112c-228">CDN の Office.js を使用しない場合、アドインで例外が表示されることがあります。**isSetSupported** が定義されていないためです。</span><span class="sxs-lookup"><span data-stu-id="e112c-228">If you don't use Office.js from the CDN, your add-in might generate exceptions because **isSetSupported** will be undefined.</span></span> <span data-ttu-id="e112c-229">詳細については、「[最新の JavaScript API for Office ライブラリを指定する](#specify-the-latest-javascript-api-for-office-library)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e112c-229">For more information, see [Specify the latest JavaScript API for Office library](#specify-the-latest-javascript-api-for-office-library).</span></span>

<span data-ttu-id="e112c-230">次のコードの例は、さまざまな要件セットや API メンバーをサポートするさまざまな Office ホストにおいて、アドインで各種の機能を提供する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="e112c-230">The following code example shows how an add-in can provide different functionality for different Office hosts that might support different requirement sets or API members.</span></span>

```js
if (Office.context.requirements.isSetSupported('WordApi', 1.1))
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


## <a name="runtime-checks-using-methods-not-in-a-requirement-set"></a><span data-ttu-id="e112c-231">要件セットにないメソッドを使用したランタイム チェック</span><span class="sxs-lookup"><span data-stu-id="e112c-231">Runtime checks using methods not in a requirement set</span></span>

<span data-ttu-id="e112c-232">API の一部のメンバーは、要件のセットに属していません。</span><span class="sxs-lookup"><span data-stu-id="e112c-232">Some API members don't belong to requirement sets.</span></span> <span data-ttu-id="e112c-233">これは [JavaScript API for Office](/office/dev/add-ins/reference/javascript-api-for-office) 名前空間 ([Outlook Mailbox API](/javascript/api/outlook) を除く `Office.` で始まるすべての名前空間) に属する API メンバーにのみ適用され、[Word JavaScript API](/office/dev/add-ins/reference/overview/word-add-ins-reference-overview) 名前空間 (`Word.` で始まるすべての名前空間)、[Excel JavaScript API](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview) 名前空間 (`Excel.` で始まるすべての名前空間) や [OneNote JavaScript API](/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference) (`OneNote.` で始まるすべての名前空間) に属する API メンバーには適用されません。</span><span class="sxs-lookup"><span data-stu-id="e112c-233">This only applies to API members that are part of the JavaScript API for Office namespace (anything under Office.), not API members that belong to the Word JavaScript API (anything in Word.) or Excel add-ins JavaScript API reference (anything in Excel.) namespaces.</span></span> <span data-ttu-id="e112c-234">要件セットに属さないメソッドにアドインが依存するとき、ランタイム チェックを利用し、メソッドが Office ホストでサポートされているかどうかを判断できます。たとえば、次のコード例のようになります。</span><span class="sxs-lookup"><span data-stu-id="e112c-234">When your add-in depends on a method that is not part of a requirement set, you can use the runtime check to determine whether the method is supported by the Office host, as shown in the following code example.</span></span> <span data-ttu-id="e112c-235">要件セットに属さないメソッドの詳細な一覧については、「[Office アドインの要件セット](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#methods-that-arent-part-of-a-requirement-set)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e112c-235">For a complete list of methods that don't belong to a requirement set, see [Office Add-in requirement sets](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#methods-that-arent-part-of-a-requirement-set).</span></span>

> [!NOTE]
> <span data-ttu-id="e112c-236">アドインのコードでのこの種のランタイム チェックは、限定的に使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="e112c-236">We recommend that you limit the use of this type of runtime check in your add-in's code.</span></span>

<span data-ttu-id="e112c-237">次のコードの例は、ホストが **document.setSelectedDataAsync** をサポートしているかどうかをチェックします。</span><span class="sxs-lookup"><span data-stu-id="e112c-237">The following code example checks whether the host supports  **document.setSelectedDataAsync**.</span></span>

```js
if (Office.context.document.setSelectedDataAsync)
{
    // Run code that uses document.setSelectedDataAsync.
}
```


## <a name="see-also"></a><span data-ttu-id="e112c-238">関連項目</span><span class="sxs-lookup"><span data-stu-id="e112c-238">See also</span></span>

- [<span data-ttu-id="e112c-239">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="e112c-239">Office Add-ins XML manifest</span></span>](add-in-manifests.md)
- [<span data-ttu-id="e112c-240">Office アドインの要件セット</span><span class="sxs-lookup"><span data-stu-id="e112c-240">Office Add-in requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="e112c-241">Word-Add-in-Get-Set-EditOpen-XML</span><span class="sxs-lookup"><span data-stu-id="e112c-241">Word-Add-in-Get-Set-EditOpen-XML</span></span>](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML)
