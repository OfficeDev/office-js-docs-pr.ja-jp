---
title: Office JavaScript API の API オープン仕様
description: ''
ms.date: 05/13/2019
localization_priority: Normal
ms.openlocfilehash: b9531200688afefa71ec74cb8d38c4116df46027
ms.sourcegitcommit: 944cbb5c6ce055f6db1833182b24d490d1dce01d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/14/2019
ms.locfileid: "34057621"
---
# <a name="api-open-specifications"></a><span data-ttu-id="1d09c-102">API オープン仕様</span><span class="sxs-lookup"><span data-stu-id="1d09c-102">API open specifications</span></span>

<span data-ttu-id="1d09c-103">Office JavaScript API オープン仕様では、Excel、Word、その他のホスト アプリケーション用に設計されている新しい JavaScript API に関する情報が提供されます。</span><span class="sxs-lookup"><span data-stu-id="1d09c-103">The Office JavaScript API open specifications provide information about new JavaScript APIs that are being designed for Excel, Word, and other host applications.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1d09c-104">API オープン仕様に記載されている機能は、初期設計やパブリック プレビューなどさまざまな開発段階にあるため、変更の対象となります。</span><span class="sxs-lookup"><span data-stu-id="1d09c-104">Features described in the API open specifications may be in various stages of development, such as early design or public preview, and are subject to change.</span></span> <span data-ttu-id="1d09c-105">API 機能が一般的に使用できるようになると、[リファレンス ドキュメント](/javascript/api/overview/office)は更新されます。</span><span class="sxs-lookup"><span data-stu-id="1d09c-105">When an API feature becomes generally available, the [reference documentation](/javascript/api/overview/office) will be updated.</span></span>

## <a name="open-specifications"></a><span data-ttu-id="1d09c-106">オープン仕様</span><span class="sxs-lookup"><span data-stu-id="1d09c-106">Open specifications</span></span>

<span data-ttu-id="1d09c-107">オープン仕様フェーズでは、今後の設計についてコミュニティから聞くことができます。</span><span class="sxs-lookup"><span data-stu-id="1d09c-107">The open specification phase allows us to hear from the community about upcoming designs.</span></span> <span data-ttu-id="1d09c-108">現時点では、公開されているオープン仕様フェーズの Api はありません。</span><span class="sxs-lookup"><span data-stu-id="1d09c-108">Currently, we have no APIs in the public open specification phase.</span></span> <span data-ttu-id="1d09c-109">このページは、利用可能になったときに新しいデザインで更新されます。</span><span class="sxs-lookup"><span data-stu-id="1d09c-109">We will update this page with new designs as they become available.</span></span>

## <a name="preview-apis"></a><span data-ttu-id="1d09c-110">プレビュー Api</span><span class="sxs-lookup"><span data-stu-id="1d09c-110">Preview APIs</span></span>

<span data-ttu-id="1d09c-111">オープン仕様プロセスの後、新しい Office JavaScript はプレビューフェーズに入ります。</span><span class="sxs-lookup"><span data-stu-id="1d09c-111">After the open specification process, new Office JavaScript enter the preview phase.</span></span> <span data-ttu-id="1d09c-112">プレビュー Api を使用するには、アドインが CDNhttps://appsforoffice.microsoft.com/lib/beta/hosted/office.js)の**ベータ版**ライブラリを参照する必要があります。また、office Insider プログラムに参加して、最新の office ビルドを取得する必要がある場合もあります。</span><span class="sxs-lookup"><span data-stu-id="1d09c-112">To use preview APIs, your add-in must reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) and you may also need to join the Office Insider program to get a recent Office build.</span></span>

<span data-ttu-id="1d09c-113">フィードバックはプレビュー Api のまま歓迎されます。</span><span class="sxs-lookup"><span data-stu-id="1d09c-113">Feedback is still welcome for preview APIs.</span></span> <span data-ttu-id="1d09c-114">この手順で設計が必ずしもファイナライズされるわけではありません。また、Api がリリースされたときには、それが高品質であることを確認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1d09c-114">The design is not necessarily finalized at this step and we want to ensure that when the APIs are released they are of high quality.</span></span> <span data-ttu-id="1d09c-115">問題が発生したか、または GitHub を通じて発生した問題を報告してください。</span><span class="sxs-lookup"><span data-stu-id="1d09c-115">Please report any issues you encounter or thoughts or have through GitHub.</span></span> <span data-ttu-id="1d09c-116">各ページには、ページの最後にレポートリンクがあります。</span><span class="sxs-lookup"><span data-stu-id="1d09c-116">Each page has a reporting link at the end of the page.</span></span>

### <a name="new-excel-javascript-apis"></a><span data-ttu-id="1d09c-117">新しい Excel JavaScript API</span><span class="sxs-lookup"><span data-stu-id="1d09c-117">New Excel JavaScript APIs</span></span>

<span data-ttu-id="1d09c-118">新しい Excel JavaScript API の設計のレビューにご参加ください。</span><span class="sxs-lookup"><span data-stu-id="1d09c-118">Join us in reviewing our design for new Excel JavaScript APIs.</span></span> <span data-ttu-id="1d09c-119">最新の更新は、[Excel JavaScript API の要件セットのページ](../requirement-sets/excel-api-requirement-sets.md#excel-javascript-preview-apis)にあります。</span><span class="sxs-lookup"><span data-stu-id="1d09c-119">The latest updates can be found in the [Excel JavaScript API requirement sets page](../requirement-sets/excel-api-requirement-sets.md#excel-javascript-preview-apis).</span></span>

<span data-ttu-id="1d09c-120">**詳細については [Excel JavaScript プレビュー API](/javascript/api/excel)を参照し、フィードバックを提供してください。**</span><span class="sxs-lookup"><span data-stu-id="1d09c-120">**See the [Excel JavaScript preview APIs](/javascript/api/excel) to learn more and provide your feedback.**</span></span>

### <a name="new-word-javascript-apis"></a><span data-ttu-id="1d09c-121">新しい Word JavaScript API</span><span class="sxs-lookup"><span data-stu-id="1d09c-121">New Word JavaScript APIs</span></span>

<span data-ttu-id="1d09c-122">新しい Word JavaScript API の設計のレビューにご参加ください。</span><span class="sxs-lookup"><span data-stu-id="1d09c-122">Join us in reviewing our design for new Word JavaScript APIs.</span></span> <span data-ttu-id="1d09c-123">最新の更新プログラムについては、「 [JAVASCRIPT API の要件セット」ページ](../requirement-sets/word-api-requirement-sets.md#word-javascript-preview-apis)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1d09c-123">The latest updates can be found in the [Word JavaScript API requirement sets page](../requirement-sets/word-api-requirement-sets.md#word-javascript-preview-apis).</span></span>

<span data-ttu-id="1d09c-124">**詳細については、 [Word JavaScript Preview api](/javascript/api/word)を参照し、フィードバックを提供してください。**</span><span class="sxs-lookup"><span data-stu-id="1d09c-124">**See the [Word JavaScript preview APIs](/javascript/api/word) to learn more and provide your feedback.**</span></span>
