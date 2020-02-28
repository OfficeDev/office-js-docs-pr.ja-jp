---
title: マニフェスト ファイルの HighResolutionIconUrl 要素
description: ''
ms.date: 12/04/2018
localization_priority: Normal
ms.openlocfilehash: 41008be6b60d260bef78808af2b8dee1fbd0864a
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325270"
---
# <a name="highresolutioniconurl-element"></a><span data-ttu-id="f417d-102">HighResolutionIconUrl 要素</span><span class="sxs-lookup"><span data-stu-id="f417d-102">HighResolutionIconUrl element</span></span>

<span data-ttu-id="f417d-103">高 DPI 画面での挿入 UX と Office ストアで Office アドインを表すために使用されるイメージの URL を指定します。</span><span class="sxs-lookup"><span data-stu-id="f417d-103">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store on high DPI screens.</span></span>

<span data-ttu-id="f417d-104">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="f417d-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="f417d-105">構文</span><span class="sxs-lookup"><span data-stu-id="f417d-105">Syntax</span></span>

```XML
<HighResolutionIconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="f417d-106">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="f417d-106">Can contain</span></span>

[<span data-ttu-id="f417d-107">Override</span><span class="sxs-lookup"><span data-stu-id="f417d-107">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="f417d-108">属性</span><span class="sxs-lookup"><span data-stu-id="f417d-108">Attributes</span></span>

|<span data-ttu-id="f417d-109">**属性**</span><span class="sxs-lookup"><span data-stu-id="f417d-109">**Attribute**</span></span>|<span data-ttu-id="f417d-110">**型**</span><span class="sxs-lookup"><span data-stu-id="f417d-110">**Type**</span></span>|<span data-ttu-id="f417d-111">**必須**</span><span class="sxs-lookup"><span data-stu-id="f417d-111">**Required**</span></span>|<span data-ttu-id="f417d-112">**説明**</span><span class="sxs-lookup"><span data-stu-id="f417d-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="f417d-113">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="f417d-113">DefaultValue</span></span>|<span data-ttu-id="f417d-114">文字列 (URL)</span><span class="sxs-lookup"><span data-stu-id="f417d-114">string (URL)</span></span>|<span data-ttu-id="f417d-115">必須</span><span class="sxs-lookup"><span data-stu-id="f417d-115">required</span></span>|<span data-ttu-id="f417d-116">この設定の既定値を指定します。この値は、[DefaultLocale](defaultlocale.md) 要素に指定されるロケールを対象としています。</span><span class="sxs-lookup"><span data-stu-id="f417d-116">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="f417d-117">注釈</span><span class="sxs-lookup"><span data-stu-id="f417d-117">Remarks</span></span>

<span data-ttu-id="f417d-118">メールアドインの場合は、[**ファイル** > の**管理**] [アドイン] UI にアイコンが表示されます。</span><span class="sxs-lookup"><span data-stu-id="f417d-118">For a mail add-in, the icon is displayed in the **File** > **Manage add-ins** UI .</span></span> <span data-ttu-id="f417d-119">コンテンツ アドインまたは作業ウィンドウ アドインでは、アイコンは、**[挿入]** > **[アドイン]** UI に表示されます。</span><span class="sxs-lookup"><span data-stu-id="f417d-119">For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span></span>

<span data-ttu-id="f417d-120">画像のファイル形式は GIF、JPG、PNG、EXIF、BMP、TIFF のいずれかにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="f417d-120">The image must be in one of the following file formats: GIF, JPG, PNG, EXIF, BMP, or TIFF.</span></span> <span data-ttu-id="f417d-121">コンテンツおよび作業ウィンドウ アプリの推奨される画像の解像度は 64 x 64 ピクセルです。</span><span class="sxs-lookup"><span data-stu-id="f417d-121">For content and task pane apps, the recommended image resolution is 64 x 64 pixels.</span></span> <span data-ttu-id="f417d-122">メール アプリの画像は 128 × 128 ピクセルにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="f417d-122">For mail apps, the image must be 128 x 128 pixels.</span></span> <span data-ttu-id="f417d-123">詳細については、「[効果的な AppSource と Office 内の登録リストを作成する](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity)」の「_アプリに一貫性のあるビジュアル ID を作成する_」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="f417d-123">For more information, see the section  _Create a consistent visual identity for your app_ in [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span></span>
