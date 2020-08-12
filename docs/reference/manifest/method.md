---
title: マニフェスト ファイルの Method 要素
description: Method 要素は、office アドインをアクティブにするために必要な、Office JavaScript API からの個別のメソッドを指定します。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 0e3e74a73a3422a7789e82d6f0e7a516bd795ca8
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641326"
---
# <a name="method-element"></a><span data-ttu-id="e4fa3-103">Method 要素</span><span class="sxs-lookup"><span data-stu-id="e4fa3-103">Method element</span></span>

<span data-ttu-id="e4fa3-104">Office JavaScript API から、Office アドインをアクティブにするために必要な個別のメソッドを指定します。</span><span class="sxs-lookup"><span data-stu-id="e4fa3-104">Specifies an individual method from the Office JavaScript API that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="e4fa3-105">**アドインの種類:** コンテンツ、作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e4fa3-105">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="e4fa3-106">構文</span><span class="sxs-lookup"><span data-stu-id="e4fa3-106">Syntax</span></span>

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a><span data-ttu-id="e4fa3-107">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="e4fa3-107">Contained in</span></span>

[<span data-ttu-id="e4fa3-108">Methods</span><span class="sxs-lookup"><span data-stu-id="e4fa3-108">Methods</span></span>](methods.md)

## <a name="attributes"></a><span data-ttu-id="e4fa3-109">属性</span><span class="sxs-lookup"><span data-stu-id="e4fa3-109">Attributes</span></span>

|<span data-ttu-id="e4fa3-110">属性</span><span class="sxs-lookup"><span data-stu-id="e4fa3-110">Attribute</span></span>|<span data-ttu-id="e4fa3-111">型</span><span class="sxs-lookup"><span data-stu-id="e4fa3-111">Type</span></span>|<span data-ttu-id="e4fa3-112">必須</span><span class="sxs-lookup"><span data-stu-id="e4fa3-112">Required</span></span>|<span data-ttu-id="e4fa3-113">説明</span><span class="sxs-lookup"><span data-stu-id="e4fa3-113">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="e4fa3-114">名前</span><span class="sxs-lookup"><span data-stu-id="e4fa3-114">Name</span></span>|<span data-ttu-id="e4fa3-115">string</span><span class="sxs-lookup"><span data-stu-id="e4fa3-115">string</span></span>|<span data-ttu-id="e4fa3-116">必須</span><span class="sxs-lookup"><span data-stu-id="e4fa3-116">required</span></span>|<span data-ttu-id="e4fa3-117">必要なメソッドの名前をその親オブジェクトで修飾して指定します。</span><span class="sxs-lookup"><span data-stu-id="e4fa3-117">Specifies the name of the required method qualified with its parent object.</span></span> <span data-ttu-id="e4fa3-118">たとえば、メソッドを指定するには、を `getSelectedDataAsync` 指定する必要があり `"Document.getSelectedDataAsync"` ます。</span><span class="sxs-lookup"><span data-stu-id="e4fa3-118">For example, to specify the `getSelectedDataAsync` method, you must specify `"Document.getSelectedDataAsync"`.</span></span>|

## <a name="remarks"></a><span data-ttu-id="e4fa3-119">注釈</span><span class="sxs-lookup"><span data-stu-id="e4fa3-119">Remarks</span></span>

<span data-ttu-id="e4fa3-120">`Methods`および要素は、 `Method` メールアドインではサポートされていません。要件セットの詳細については、「 [Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e4fa3-120">The `Methods` and `Method` elements aren't supported by mail add-ins. For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="e4fa3-121">個々のメソッドの最小バージョン要件を指定する方法がないため、メソッドが実行時に使用可能であることを確認するには、そのメソッドをアドインのスクリプトで呼び出す際に、**if** ステートメントも使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e4fa3-121">Because there is no way to specify the minimum version requirement for individual methods, to make sure that a method is available at runtime, you should also use an **if** statement when calling that method in the script of your add-in.</span></span> <span data-ttu-id="e4fa3-122">これを行う方法の詳細については、「 [Office JAVASCRIPT API に](../../develop/understanding-the-javascript-api-for-office.md)ついて」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e4fa3-122">For more information about how to do this, see [Understanding the Office JavaScript API](../../develop/understanding-the-javascript-api-for-office.md).</span></span>
