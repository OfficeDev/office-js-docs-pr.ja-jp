---
title: JavaScript API for Office
description: ''
ms.date: 10/09/2018
localization_priority: Priority
ms.openlocfilehash: 51ae2a2391a63ab20a33e8ae590834f33311230a
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/23/2019
ms.locfileid: "29387843"
---
# <a name="javascript-api-for-office"></a><span data-ttu-id="2c310-102">JavaScript API for Office</span><span class="sxs-lookup"><span data-stu-id="2c310-102">JavaScript API for Office</span></span>

<span data-ttu-id="2c310-p101">JavaScript API for Office を使用すると、Office ホスト アプリケーションのオブジェクト モデルと対話する Web アプリケーションを作成できます。ユーザーのアプリケーションは、スクリプト ローダーである office.js ライブラリを参照します。Office.js ライブラリは、アドインを実行している Office アプリケーションに適用可能なオブジェクト モデルを読み込みます。次の JavaScript オブジェクト モデルを使用できます。</span><span class="sxs-lookup"><span data-stu-id="2c310-p101">The JavaScript API for Office enables you to create web applications that interact with the object models in Office host applications. Your application will reference the office.js library, which is a script loader. The office.js library loads the object models that are applicable to the Office application that is running the add-in. You can use the following JavaScript object models:</span></span>

- <span data-ttu-id="2c310-107">**共通 API** - **Office 2013** で導入された API。</span><span class="sxs-lookup"><span data-stu-id="2c310-107">**Common APIs** - APIs that were introduced with **Office 2013**.</span></span> <span data-ttu-id="2c310-108">これは、**すべての Office ホスト アプリケーション**に読み込まれ、アドイン アプリケーションを Office クライアント アプリケーションに接続します。</span><span class="sxs-lookup"><span data-stu-id="2c310-108">This is loaded for **all Office host applications** and connects your add-in application with the Office client application.</span></span> <span data-ttu-id="2c310-109">オブジェクト モデルには、Office クライアントに固有の API と複数の Office クライアントのホスト アプリケーションに適用可能な API が含まれています。</span><span class="sxs-lookup"><span data-stu-id="2c310-109">The object model contains APIs that are specific to Office clients, and APIs that are applicable to multiple Office client host applications.</span></span> <span data-ttu-id="2c310-110">このコンテンツは、すべて**共通 API** の下にあります。</span><span class="sxs-lookup"><span data-stu-id="2c310-110">All of this content is under **Common API**.</span></span> <span data-ttu-id="2c310-111">このオブジェクト モデルは、コールバックを使用します。</span><span class="sxs-lookup"><span data-stu-id="2c310-111">This object model uses callbacks.</span></span> 

  <span data-ttu-id="2c310-112">**Outlook** でも共通 API 構文が使用されます。</span><span class="sxs-lookup"><span data-stu-id="2c310-112">**Outlook** also uses the Common API syntax.</span></span> <span data-ttu-id="2c310-113">Office というエイリアスの下にあるすべてのものの中には、Office アドインから Office ドキュメント、ワークシート、プレゼンテーション、メール アイテム、プロジェクトのコンテンツを操作するスクリプトの記述に利用できるオブジェクトが含まれています。アドインが Office 2013 以降を対象としている場合には、これらの共通 API を使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="2c310-113">Everything under the alias Office contains objects you can use to write scripts that interact with content in Office documents, worksheets, presentations, mail items, and projects from your Office Add-ins. You must use these Common APIs if your add-in will target Office 2013 and later.</span></span> <span data-ttu-id="2c310-114">このオブジェクト モデルは、コールバックを使用します。</span><span class="sxs-lookup"><span data-stu-id="2c310-114">This object model uses callbacks.</span></span>

- <span data-ttu-id="2c310-115">**ホスト固有 API** - **Office 2016** で導入された API。</span><span class="sxs-lookup"><span data-stu-id="2c310-115">**Host-specific APIs** - APIs that were introduced with **Office 2016**.</span></span> <span data-ttu-id="2c310-116">このオブジェクト モデルは、Office クライアントの使用時に見られる使い慣れたオブジェクトに対応するホスト固有の厳密に型指定されたオブジェクトを提供し、Office JavaScript API の将来像を表すものです。</span><span class="sxs-lookup"><span data-stu-id="2c310-116">This object model provides host-specific strongly-typed objects that correspond to familiar objects that you see when you use Office clients, and represents the future of Office JavaScript APIs.</span></span> <span data-ttu-id="2c310-117">現在、ホスト固有の API には、Word JavaScript API と Excel JavaScript API が含まれています。</span><span class="sxs-lookup"><span data-stu-id="2c310-117">The host-specific APIs currently include the Word JavaScript API and the Excel JavaScript API.</span></span>

## <a name="supported-host-applications"></a><span data-ttu-id="2c310-118">サポートされるホスト アプリケーション</span><span class="sxs-lookup"><span data-stu-id="2c310-118">Supported host applications</span></span>

- [<span data-ttu-id="2c310-119">Excel</span><span class="sxs-lookup"><span data-stu-id="2c310-119">Excel</span></span>](overview/excel-add-ins-reference-overview.md)
- [<span data-ttu-id="2c310-120">OneNote</span><span class="sxs-lookup"><span data-stu-id="2c310-120">OneNote</span></span>](overview/onenote-add-ins-javascript-reference.md)
- [<span data-ttu-id="2c310-121">Outlook</span><span class="sxs-lookup"><span data-stu-id="2c310-121">Outlook</span></span>](requirement-sets/outlook-api-requirement-sets.md)
- [<span data-ttu-id="2c310-122">Visio</span><span class="sxs-lookup"><span data-stu-id="2c310-122">Visio</span></span>](overview/visio-javascript-reference-overview.md)
- [<span data-ttu-id="2c310-123">Word</span><span class="sxs-lookup"><span data-stu-id="2c310-123">Word</span></span>](overview/word-add-ins-reference-overview.md)
- [<span data-ttu-id="2c310-124">共通 API</span><span class="sxs-lookup"><span data-stu-id="2c310-124">Common API</span></span>](requirement-sets/office-add-in-requirement-sets.md)

> [!NOTE] 
> <span data-ttu-id="2c310-125">[PowerPoint と Project](requirement-sets/powerpoint-and-project-note.md) は JavaScript API で作成されたアドインをサポートしています。</span><span class="sxs-lookup"><span data-stu-id="2c310-125">[PowerPoint and Project](requirement-sets/powerpoint-and-project-note.md) support add-ins made with the JavaScript API.</span></span> <span data-ttu-id="2c310-126">ただし、現在はホスト固有の API は含まれていません。</span><span class="sxs-lookup"><span data-stu-id="2c310-126">However, they currently do not have host-specific APIs.</span></span> <span data-ttu-id="2c310-127">これらのホストとは共通 API を通じて対話します。</span><span class="sxs-lookup"><span data-stu-id="2c310-127">You interact with these hosts through the Common API.</span></span>

<span data-ttu-id="2c310-128">[サポートされるホストとその他の要件](../concepts/requirements-for-running-office-add-ins.md)の詳細について説明します。</span><span class="sxs-lookup"><span data-stu-id="2c310-128">Learn more about [supported hosts and other requirements](../concepts/requirements-for-running-office-add-ins.md).</span></span>

## <a name="open-api-specifications"></a><span data-ttu-id="2c310-129">Open API の仕様</span><span class="sxs-lookup"><span data-stu-id="2c310-129">Open API specifications</span></span>

<span data-ttu-id="2c310-p106">新しい Office アドイン用の API の設計と開発にあたり、[Open API の仕様](openspec.md) ページでこれらに対するフィードバックの提供が可能になります。パイプラインの新機能をご確認いただき、設計の仕様に関する情報をお寄せください。</span><span class="sxs-lookup"><span data-stu-id="2c310-p106">As we design and develop new APIs for Office Add-ins, we'll make them available for your feedback on our [Open API specifications](openspec.md) page. Find out what new features are in the pipeline, and provide your input on our design specifications.</span></span>

## <a name="see-also"></a><span data-ttu-id="2c310-132">関連項目</span><span class="sxs-lookup"><span data-stu-id="2c310-132">See also</span></span>

- [<span data-ttu-id="2c310-133">Office JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="2c310-133">Office JavaScript API reference</span></span>](https://docs.microsoft.com/javascript/api/overview/office)
