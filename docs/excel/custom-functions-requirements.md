---
ms.date: 01/30/2019
description: さまざまなプラットフォームでカスタム関数を使用する方法を説明します。
title: カスタム関数の要件 (プレビュー)
localization_priority: Priority
ms.openlocfilehash: 0226c5e129794e8105966e753e791ec20766e10c
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449247"
---
# <a name="custom-functions-requirements"></a><span data-ttu-id="b6b6a-103">カスタム関数の要件</span><span class="sxs-lookup"><span data-stu-id="b6b6a-103">Custom functions requirements</span></span>

<span data-ttu-id="b6b6a-104">カスタム関数は、現在、次のプラットフォームで開発者向けのプレビューとして利用できます。</span><span class="sxs-lookup"><span data-stu-id="b6b6a-104">Custom functions are currently available in developer preview on the following platforms:</span></span>

- <span data-ttu-id="b6b6a-105">Excel Online</span><span class="sxs-lookup"><span data-stu-id="b6b6a-105">Excel Online</span></span>
- <span data-ttu-id="b6b6a-106">Windows 版 Excel (64 ビット バージョン 1810 以降)。</span><span class="sxs-lookup"><span data-stu-id="b6b6a-106">Excel for Windows (64-bit version 1810 or later).</span></span> <span data-ttu-id="b6b6a-107">現時点で、Windows 版 Excel 32 ビットではすべてのシナリオが使用できるとは限りません。</span><span class="sxs-lookup"><span data-stu-id="b6b6a-107">At present, Excel for Windows 32-bit may not work for all scenarios.</span></span> 
- <span data-ttu-id="b6b6a-108">Excel for Mac (バージョン 13.329 以降)</span><span class="sxs-lookup"><span data-stu-id="b6b6a-108">Excel for Mac (version 13.329 or later)</span></span>

<span data-ttu-id="b6b6a-109">カスタム関数機能は現在プレビュー段階であり、変更される可能性があることをご留意ください。</span><span class="sxs-lookup"><span data-stu-id="b6b6a-109">Note that the custom functions feature is currently in preview and subject to change.</span></span> <span data-ttu-id="b6b6a-110">運用環境での使用は、現在サポートされていません。</span><span class="sxs-lookup"><span data-stu-id="b6b6a-110">It is not currently supported for use in production environments.</span></span>

## <a name="excel-online"></a><span data-ttu-id="b6b6a-111">Excel Online</span><span class="sxs-lookup"><span data-stu-id="b6b6a-111">Excel Online</span></span>
<span data-ttu-id="b6b6a-112">Excel Online でカスタム関数を使用するには、Office 365 サブスクリプションまたは [Microsoft アカウント](https://account.microsoft.com/account)のいずれかを使用してログインします。</span><span class="sxs-lookup"><span data-stu-id="b6b6a-112">To use custom functions within Excel Online, login by using either your Office 365 subscription or a [Microsoft account](https://account.microsoft.com/account).</span></span> 

## <a name="excel-for-windows-and-excel-for-mac"></a><span data-ttu-id="b6b6a-113">Excel for Windows および Excel for Mac</span><span class="sxs-lookup"><span data-stu-id="b6b6a-113">Excel for Windows and Excel for Mac</span></span>
<span data-ttu-id="b6b6a-114">Windows 版 Excel または Excel for Mac 内でカスタム関数を使用するには、Office 365 サブスクリプションがあり、[Office Insider](https://products.office.com/office-insider) プログラム (**Insider** レベル -- 旧称 "Insider Fast") に参加しており、この前述したように十分に新しいバージョンの Excel を使用している必要があります。</span><span class="sxs-lookup"><span data-stu-id="b6b6a-114">To use custom functions within Excel for Windows or Excel for Mac, you must have an Office 365 subscription, join the [Office Insider](https://products.office.com/office-insider) program (**Insider** level -- formerly called "Insider Fast"), and use a sufficiently recent build of Excel (as specified previously).</span></span>

<span data-ttu-id="b6b6a-115">Windows ストア からダウンロードしたバージョンのデスクトップ版 Office を使用している場合は、カスタム関数を使用するには、[Windows Insider](https://insider.windows.com/) プログラムに **Insider** レベル (旧称 "Insider Fast") で参加し、2018 年 4 月以降の更新プログラムのバージョンを実行している必要があります。</span><span class="sxs-lookup"><span data-stu-id="b6b6a-115">If you are using a version of Office on your desktop which you downloaded from the Windows Store, you must be part of the [Windows Insider](https://insider.windows.com/) program at the **Insider** level (formerly called "Insider Fast"), running the April 2018 Update version or later to use custom functions.</span></span> <span data-ttu-id="b6b6a-116">これは、2019 年 1 月時点での新しい変更点です。</span><span class="sxs-lookup"><span data-stu-id="b6b6a-116">This is a new change as of January 2019.</span></span>

## <a name="subscribe-to-office-365"></a><span data-ttu-id="b6b6a-117">Office 365 のサブスクリプション</span><span class="sxs-lookup"><span data-stu-id="b6b6a-117">Subscribe to Office 365</span></span>
<span data-ttu-id="b6b6a-118">Office 365 サブスクリプションをまだお持ちでない場合は、[Office 365 Developer Program](https://developer.microsoft.com/ja-JP/office/dev-program) に参加することで入手できます。</span><span class="sxs-lookup"><span data-stu-id="b6b6a-118">If you don't already have an Office 365 subscription, you can get one by joining the [Office 365 Developer Program](https://developer.microsoft.com/ja-JP/office/dev-program).</span></span>


* [<span data-ttu-id="b6b6a-119">カスタム関数の概要</span><span class="sxs-lookup"><span data-stu-id="b6b6a-119">Custom functions overview</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="b6b6a-120">カスタム関数の変更ログ</span><span class="sxs-lookup"><span data-stu-id="b6b6a-120">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="b6b6a-121">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="b6b6a-121">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="b6b6a-122">チュートリアル: Excel でカスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="b6b6a-122">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
