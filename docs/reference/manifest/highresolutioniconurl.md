---
title: マニフェスト ファイルの HighResolutionIconUrl 要素
description: ''
ms.date: 12/04/2018
localization_priority: Normal
ms.openlocfilehash: 5264fc969bda30a9b2212996800b984533a3188c
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452089"
---
# <a name="highresolutioniconurl-element"></a><span data-ttu-id="bd0bb-102">HighResolutionIconUrl 要素</span><span class="sxs-lookup"><span data-stu-id="bd0bb-102">HighResolutionIconUrl element</span></span>

<span data-ttu-id="bd0bb-103">高 DPI 画面での挿入 UX と Office ストアで Office アドインを表すために使用されるイメージの URL を指定します。</span><span class="sxs-lookup"><span data-stu-id="bd0bb-103">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store on high DPI screens.</span></span>

<span data-ttu-id="bd0bb-104">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="bd0bb-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="bd0bb-105">構文</span><span class="sxs-lookup"><span data-stu-id="bd0bb-105">Syntax</span></span>

```XML
<HighResolutionIconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="bd0bb-106">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="bd0bb-106">Can contain</span></span>

[<span data-ttu-id="bd0bb-107">Override</span><span class="sxs-lookup"><span data-stu-id="bd0bb-107">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="bd0bb-108">属性</span><span class="sxs-lookup"><span data-stu-id="bd0bb-108">Attributes</span></span>

|<span data-ttu-id="bd0bb-109">**属性**</span><span class="sxs-lookup"><span data-stu-id="bd0bb-109">**Attribute**</span></span>|<span data-ttu-id="bd0bb-110">**型**</span><span class="sxs-lookup"><span data-stu-id="bd0bb-110">**Type**</span></span>|<span data-ttu-id="bd0bb-111">**必須**</span><span class="sxs-lookup"><span data-stu-id="bd0bb-111">**Required**</span></span>|<span data-ttu-id="bd0bb-112">**説明**</span><span class="sxs-lookup"><span data-stu-id="bd0bb-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="bd0bb-113">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="bd0bb-113">DefaultValue</span></span>|<span data-ttu-id="bd0bb-114">文字列 (URL)</span><span class="sxs-lookup"><span data-stu-id="bd0bb-114">string (URL)</span></span>|<span data-ttu-id="bd0bb-115">必須</span><span class="sxs-lookup"><span data-stu-id="bd0bb-115">required</span></span>|<span data-ttu-id="bd0bb-116">この設定の既定値を指定します。この値は、[DefaultLocale](defaultlocale.md) 要素に指定されるロケールを対象としています。</span><span class="sxs-lookup"><span data-stu-id="bd0bb-116">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="bd0bb-117">注釈</span><span class="sxs-lookup"><span data-stu-id="bd0bb-117">Remarks</span></span>

<span data-ttu-id="bd0bb-p101">メール アドインの場合、アイコンは、**[ファイル]**  >  **[アドインの管理]** UI に表示されます。コンテンツ アドインまたは作業ウィンドウ アドインでは、アイコンは、**[挿入]**  >  **[アドイン]** UI に表示されます。</span><span class="sxs-lookup"><span data-stu-id="bd0bb-p101">For a mail add-in, the icon is displayed in the  **File** > **Manage add-ins** UI . For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span></span>

<span data-ttu-id="bd0bb-120">画像のファイル形式は GIF、JPG、PNG、EXIF、BMP、TIFF のいずれかにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="bd0bb-120">The image must be in one of the following file formats: GIF, JPG, PNG, EXIF, BMP, or TIFF.</span></span> <span data-ttu-id="bd0bb-121">コンテンツおよび作業ウィンドウ アプリの推奨される画像の解像度は 64 x 64 ピクセルです。</span><span class="sxs-lookup"><span data-stu-id="bd0bb-121">For content and task pane apps, the recommended image resolution is 64 x 64 pixels.</span></span> <span data-ttu-id="bd0bb-122">メール アプリの画像は 128 × 128 ピクセルにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="bd0bb-122">For mail apps, the image must be 128 x 128 pixels.</span></span> <span data-ttu-id="bd0bb-123">詳細については、「[効果的な AppSource と Office 内の登録リストを作成する](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity)」の「_アプリに一貫性のあるビジュアル ID を作成する_」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="bd0bb-123">For more information, see the section  _Create a consistent visual identity for your app_ in [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span></span>
