---
ms.date: 01/08/2019
description: Excel のカスタム関数に対する最新の更新内容を確認します。
title: カスタム関数の変更ログ (プレビュー)
localization_priority: Normal
ms.openlocfilehash: 03e4dd922ac3895e11a508f97e7ac3fa3e7b1cb0
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449276"
---
# <a name="custom-functions-changelog-preview"></a><span data-ttu-id="e9548-103">カスタム関数の変更ログ (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="e9548-103">Custom functions changelog (preview)</span></span>

<span data-ttu-id="e9548-104">Excel カスタム関数は現時点で引き続きプレビュー段階です。つまり、変更点や新しい関数のリリースなど本製品に対して変更が頻繁に生じています。</span><span class="sxs-lookup"><span data-stu-id="e9548-104">Excel custom functions is still in preview and that means there are frequent changes to the product, including changes and the release of new features.</span></span> <span data-ttu-id="e9548-105">この変更ログでは、本製品に対して加えられた変更に関する最新情報を取り上げます。</span><span class="sxs-lookup"><span data-stu-id="e9548-105">This changelog provides the most up-to-date information about any changes to the product.</span></span>

- <span data-ttu-id="e9548-106">**2017 年 11 月 7 日**: カスタム関数のプレビューとサンプルを公開\*</span><span class="sxs-lookup"><span data-stu-id="e9548-106">**Nov 7, 2017**: Shipped\* the custom functions preview and samples</span></span>
- <span data-ttu-id="e9548-107">**2017 年 11 月 20 日**: ビルド 8801 以降を使用する場合の互換性バグを修正</span><span class="sxs-lookup"><span data-stu-id="e9548-107">**Nov 20, 2017**: Fixed compatibility bug for those using builds 8801 and later</span></span>
- <span data-ttu-id="e9548-108">**2017 年 11 月 28 日**: 非同期関数のキャンセルのサポートを公開\* (ストリーミング機能の変更が必要)</span><span class="sxs-lookup"><span data-stu-id="e9548-108">**Nov 28, 2017**: Shipped\* support for cancellation on asynchronous functions (requires change for streaming functions)</span></span>
- <span data-ttu-id="e9548-109">**2018 年 5 月 7 日**: Mac、Excel Online、およびインプロセスで実行される同期関数へのサポートを公開\*</span><span class="sxs-lookup"><span data-stu-id="e9548-109">**May 7, 2018**: Shipped\* support for Mac, Excel Online, and synchronous functions running in-process</span></span>
- <span data-ttu-id="e9548-110">**2018 年 9 月 20日**: JavaScript ランタイムのカスタム関数へのサポートを公開。</span><span class="sxs-lookup"><span data-stu-id="e9548-110">**September 20, 2018**: Shipped support for custom functions JavaScript runtime.</span></span> <span data-ttu-id="e9548-111">詳細については、「[Excel カスタム関数のランタイム](custom-functions-runtime.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="e9548-111">For more information, see [Runtime for Excel custom functions](custom-functions-runtime.md).</span></span>
- <span data-ttu-id="e9548-112">**2018 年 10 月 20 日**: [10 月の Insider ビルド](https://support.office.com/en-us/article/what-s-new-for-office-insiders-c152d1e2-96ff-4ce9-8c14-e74e13847a24)では、カスタム関数は、 Windows デスクトップ用およびオンライン用の[カスタム定義メタデータ](custom-functions-json.md)で 'id' パラメーターが必要になりました。</span><span class="sxs-lookup"><span data-stu-id="e9548-112">**October 20, 2018**: With the [October Insiders build](https://support.office.com/en-us/article/what-s-new-for-office-insiders-c152d1e2-96ff-4ce9-8c14-e74e13847a24), custom functions now require the 'id' parameter in your [custom functions metadata](custom-functions-json.md) for Windows Desktop and Online.</span></span> <span data-ttu-id="e9548-113">Mac では、このパラメーターは無視します。</span><span class="sxs-lookup"><span data-stu-id="e9548-113">On Mac, this parameter should be ignored.</span></span> <span data-ttu-id="e9548-114">カスタム関数では、オプションのパラメーターおよび `any` 戻り値の型もサポートされるようになりました。</span><span class="sxs-lookup"><span data-stu-id="e9548-114">Custom functions also now support optional parameters and the `any` return type.</span></span>
- <span data-ttu-id="e9548-115">**2018 年 12 月 12 日**: カスタム関数にセル アドレスを検索する手段が備わりました。</span><span class="sxs-lookup"><span data-stu-id="e9548-115">**December 12, 2018**: Custom functions now include a way to discover a cell's address.</span></span> <span data-ttu-id="e9548-116">詳しくは、「[カスタム関数が呼び出したセルを特定する](custom-functions-overview.md#determine-which-cell-invoked-your-custom-function)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="e9548-116">For more information, see [Determine which cell invoked your custom function](custom-functions-overview.md#determine-which-cell-invoked-your-custom-function).</span></span>
- <span data-ttu-id="e9548-117">**2019 年 1 月 8 日**: バインド メソッド `CustomFunctionMapping()` が `CustomFunctions.associate()` に変更されました。</span><span class="sxs-lookup"><span data-stu-id="e9548-117">**January 8, 2019**: Binding method `CustomFunctionMapping()` has been altered to `CustomFunctions.associate()`.</span></span> <span data-ttu-id="e9548-118">詳細については、「[カスタム関数のベスト プラクティス](custom-functions-best-practices.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e9548-118">For more information, see [Custom functions best practices (preview)](custom-functions-best-practices.md).</span></span>

<span data-ttu-id="e9548-119">\* [Office Insider](https://products.office.com/office-insider) チャンネル (旧称 "Insider Fast") に対して</span><span class="sxs-lookup"><span data-stu-id="e9548-119">\* to the [Office Insider](https://products.office.com/office-insider) channel (formerly called "Insider Fast")</span></span>

<span data-ttu-id="e9548-120">製品の既知の問題の一覧については、「[既知の問題](custom-functions-overview.md#known-issues)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="e9548-120">For a list of known issues with the product, see [Known Issues](custom-functions-overview.md#known-issues).</span></span> 

## <a name="see-also"></a><span data-ttu-id="e9548-121">関連項目</span><span class="sxs-lookup"><span data-stu-id="e9548-121">See also</span></span>

* [<span data-ttu-id="e9548-122">カスタム関数の概要</span><span class="sxs-lookup"><span data-stu-id="e9548-122">Custom functions overview</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="e9548-123">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="e9548-123">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="e9548-124">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="e9548-124">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="e9548-125">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="e9548-125">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="e9548-126">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="e9548-126">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="e9548-127">カスタム関数のデバッグ</span><span class="sxs-lookup"><span data-stu-id="e9548-127">Custom functions debugging</span></span>](custom-functions-debugging.md)