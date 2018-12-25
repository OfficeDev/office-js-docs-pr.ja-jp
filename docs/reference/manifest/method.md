---
title: マニフェスト ファイルの Method 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: fded84344182bb45597b00a794f18defaa44d3b3
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432824"
---
# <a name="method-element"></a><span data-ttu-id="68b68-102">Method 要素</span><span class="sxs-lookup"><span data-stu-id="68b68-102">Method element</span></span>

<span data-ttu-id="68b68-103">Office アドインをアクティブにするために必要な JavaScript API for Office の個別のメソッドを指定します。</span><span class="sxs-lookup"><span data-stu-id="68b68-103">Specifies an individual method from the JavaScript API for Office that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="68b68-104">**アドインの種類:** コンテンツ、作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="68b68-104">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="68b68-105">構文</span><span class="sxs-lookup"><span data-stu-id="68b68-105">Syntax</span></span>

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a><span data-ttu-id="68b68-106">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="68b68-106">Contained in</span></span>

[<span data-ttu-id="68b68-107">Methods</span><span class="sxs-lookup"><span data-stu-id="68b68-107">Methods</span></span>](methods.md)

## <a name="attributes"></a><span data-ttu-id="68b68-108">属性</span><span class="sxs-lookup"><span data-stu-id="68b68-108">Attributes</span></span>

|<span data-ttu-id="68b68-109">**属性**</span><span class="sxs-lookup"><span data-stu-id="68b68-109">**Attribute**</span></span>|<span data-ttu-id="68b68-110">**型**</span><span class="sxs-lookup"><span data-stu-id="68b68-110">**Type**</span></span>|<span data-ttu-id="68b68-111">**必須**</span><span class="sxs-lookup"><span data-stu-id="68b68-111">**Required**</span></span>|<span data-ttu-id="68b68-112">**説明**</span><span class="sxs-lookup"><span data-stu-id="68b68-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="68b68-113">名前</span><span class="sxs-lookup"><span data-stu-id="68b68-113">Name</span></span>|<span data-ttu-id="68b68-114">string</span><span class="sxs-lookup"><span data-stu-id="68b68-114">string</span></span>|<span data-ttu-id="68b68-115">必須</span><span class="sxs-lookup"><span data-stu-id="68b68-115">required</span></span>|<span data-ttu-id="68b68-p101">必要なメソッドの名前をその親オブジェクトで修飾して指定します。たとえば、**getSelectedDataAsync** メソッドを指定するには、`"Document.getSelectedDataAsync"` と指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="68b68-p101">Specifies the name of the required method qualified with its parent object. For example, to specify the  **getSelectedDataAsync** method, you must specify `"Document.getSelectedDataAsync"`.</span></span>|

## <a name="remarks"></a><span data-ttu-id="68b68-118">解説</span><span class="sxs-lookup"><span data-stu-id="68b68-118">Remarks</span></span>

<span data-ttu-id="68b68-119">**Methods** と **Method** 要素はメール アドインではサポートされていません。要件セットの詳細については、「[Office のバージョンと要件セット](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="68b68-119">The  Methods and Method elements aren't supported by mail add-ins. For more information about requirement sets, see Specify Office hosts and API requirements.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="68b68-120">個々のメソッドの最小バージョン要件を指定する方法がないため、メソッドが実行時に使用可能であることを確認するには、そのメソッドをアドインのスクリプトで呼び出す際に、**if** ステートメントも使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="68b68-120">Important  Because there is no way to specify the minimum version requirement for individual methods, to make sure that a method is available at runtime, you should also use an  **if** statement when calling that method in the script of your add-in. For more information about how to do this, see Understanding the JavaScript API for Office.</span></span> <span data-ttu-id="68b68-121">これを行う方法の詳細については、「[JavaScript API for Office について](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="68b68-121">For more information about how to do this, see [Understanding the JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span></span>

