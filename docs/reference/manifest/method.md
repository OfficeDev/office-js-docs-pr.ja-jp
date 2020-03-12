---
title: マニフェスト ファイルの Method 要素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 74b7a8b3d0f8511d21eb0df150500850e8b93fe9
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596894"
---
# <a name="method-element"></a><span data-ttu-id="b7304-102">Method 要素</span><span class="sxs-lookup"><span data-stu-id="b7304-102">Method element</span></span>

<span data-ttu-id="b7304-103">Office JavaScript API から、Office アドインをアクティブにするために必要な個別のメソッドを指定します。</span><span class="sxs-lookup"><span data-stu-id="b7304-103">Specifies an individual method from the Office JavaScript API that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="b7304-104">**アドインの種類:** コンテンツ、作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="b7304-104">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="b7304-105">構文</span><span class="sxs-lookup"><span data-stu-id="b7304-105">Syntax</span></span>

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a><span data-ttu-id="b7304-106">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="b7304-106">Contained in</span></span>

[<span data-ttu-id="b7304-107">Methods</span><span class="sxs-lookup"><span data-stu-id="b7304-107">Methods</span></span>](methods.md)

## <a name="attributes"></a><span data-ttu-id="b7304-108">属性</span><span class="sxs-lookup"><span data-stu-id="b7304-108">Attributes</span></span>

|<span data-ttu-id="b7304-109">**属性**</span><span class="sxs-lookup"><span data-stu-id="b7304-109">**Attribute**</span></span>|<span data-ttu-id="b7304-110">**型**</span><span class="sxs-lookup"><span data-stu-id="b7304-110">**Type**</span></span>|<span data-ttu-id="b7304-111">**必須**</span><span class="sxs-lookup"><span data-stu-id="b7304-111">**Required**</span></span>|<span data-ttu-id="b7304-112">**説明**</span><span class="sxs-lookup"><span data-stu-id="b7304-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="b7304-113">名前</span><span class="sxs-lookup"><span data-stu-id="b7304-113">Name</span></span>|<span data-ttu-id="b7304-114">string</span><span class="sxs-lookup"><span data-stu-id="b7304-114">string</span></span>|<span data-ttu-id="b7304-115">必須</span><span class="sxs-lookup"><span data-stu-id="b7304-115">required</span></span>|<span data-ttu-id="b7304-116">必要なメソッドの名前をその親オブジェクトで修飾して指定します。</span><span class="sxs-lookup"><span data-stu-id="b7304-116">Specifies the name of the required method qualified with its parent object.</span></span> <span data-ttu-id="b7304-117">たとえば、 `getSelectedDataAsync`メソッドを指定するには、を指定`"Document.getSelectedDataAsync"`する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b7304-117">For example, to specify the `getSelectedDataAsync` method, you must specify `"Document.getSelectedDataAsync"`.</span></span>|

## <a name="remarks"></a><span data-ttu-id="b7304-118">注釈</span><span class="sxs-lookup"><span data-stu-id="b7304-118">Remarks</span></span>

<span data-ttu-id="b7304-119">および`Methods` `Method`要素は、メールアドインではサポートされていません。要件セットの詳細については、「 [Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b7304-119">The `Methods` and `Method` elements aren't supported by mail add-ins. For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b7304-120">個々のメソッドの最小バージョン要件を指定する方法がないため、メソッドが実行時に使用可能であることを確認するには、そのメソッドをアドインのスクリプトで呼び出す際に、**if** ステートメントも使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b7304-120">Because there is no way to specify the minimum version requirement for individual methods, to make sure that a method is available at runtime, you should also use an **if** statement when calling that method in the script of your add-in.</span></span> <span data-ttu-id="b7304-121">これを行う方法の詳細については、「 [Office JAVASCRIPT API に](../../develop/understanding-the-javascript-api-for-office.md)ついて」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b7304-121">For more information about how to do this, see [Understanding the Office JavaScript API](../../develop/understanding-the-javascript-api-for-office.md).</span></span>
