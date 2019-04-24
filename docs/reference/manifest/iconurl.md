---
title: マニフェスト ファイルの IconUrl 要素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: f7eda7ec9e4c5da8ad0b19e5e10649696d4e85c1
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452110"
---
# <a name="iconurl-element"></a><span data-ttu-id="6f2ee-102">IconUrl 要素</span><span class="sxs-lookup"><span data-stu-id="6f2ee-102">IconUrl element</span></span>

<span data-ttu-id="6f2ee-103">挿入 UX と Office ストアで Office アドインを表すために使用されるイメージの URL を指定します。</span><span class="sxs-lookup"><span data-stu-id="6f2ee-103">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store.</span></span>

<span data-ttu-id="6f2ee-104">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="6f2ee-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="6f2ee-105">構文</span><span class="sxs-lookup"><span data-stu-id="6f2ee-105">Syntax</span></span>

```XML
<IconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="6f2ee-106">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="6f2ee-106">Can contain</span></span>

[<span data-ttu-id="6f2ee-107">Override</span><span class="sxs-lookup"><span data-stu-id="6f2ee-107">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="6f2ee-108">属性</span><span class="sxs-lookup"><span data-stu-id="6f2ee-108">Attributes</span></span>

|<span data-ttu-id="6f2ee-109">**属性**</span><span class="sxs-lookup"><span data-stu-id="6f2ee-109">**Attribute**</span></span>|<span data-ttu-id="6f2ee-110">**型**</span><span class="sxs-lookup"><span data-stu-id="6f2ee-110">**Type**</span></span>|<span data-ttu-id="6f2ee-111">**必須**</span><span class="sxs-lookup"><span data-stu-id="6f2ee-111">**Required**</span></span>|<span data-ttu-id="6f2ee-112">**説明**</span><span class="sxs-lookup"><span data-stu-id="6f2ee-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="6f2ee-113">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="6f2ee-113">DefaultValue</span></span>|<span data-ttu-id="6f2ee-114">文字列</span><span class="sxs-lookup"><span data-stu-id="6f2ee-114">string</span></span>|<span data-ttu-id="6f2ee-115">必須</span><span class="sxs-lookup"><span data-stu-id="6f2ee-115">required</span></span>|<span data-ttu-id="6f2ee-116">この設定の既定値を指定します。この値は、[DefaultLocale](defaultlocale.md) 要素に指定されるロケールを対象としています。</span><span class="sxs-lookup"><span data-stu-id="6f2ee-116">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="6f2ee-117">注釈</span><span class="sxs-lookup"><span data-stu-id="6f2ee-117">Remarks</span></span>

<span data-ttu-id="6f2ee-p101">メール アドインの場合、アイコンは、**[ファイル]**  >  **[アドインの管理]** UI (Outlook) または **[設定]**  >  **[アドインの管理]** UI (Outlook Web App) に表示されます。コンテンツ アドインまたは作業ウィンドウ アドインでは、アイコンは、**[挿入]**  >  **[アドイン]** UI に表示されます。どのアドインの種類についても、アドインを Office ストアに公開すると、アイコンは Office ストア サイトでも使用されます。</span><span class="sxs-lookup"><span data-stu-id="6f2ee-p101">For a mail add-in, the icon is displayed in the  **File** > **Manage add-ins** UI (Outlook) or **Settings** > **Manage add-ins** UI (Outlook Web App). For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI. For all add-in types, the icon is also used on the Office Store site, if you publish your add-in to the Office Store.</span></span>

<span data-ttu-id="6f2ee-121">画像のファイル形式は GIF、JPG、PNG、EXIF、BMP、TIFF のいずれかにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="6f2ee-121">The image must be in one of the following file formats: GIF, JPG, PNG, EXIF, BMP, or TIFF.</span></span> <span data-ttu-id="6f2ee-122">コンテンツ アプリおよび作業ウィンドウ アプリの場合、指定する画像は 32 x 32 ピクセルにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="6f2ee-122">For content and task pane apps, the image specified must be 32 x 32 pixels.</span></span> <span data-ttu-id="6f2ee-123">メール アプリの場合、推奨される画像の解像度は 64 x 64 ピクセルです。</span><span class="sxs-lookup"><span data-stu-id="6f2ee-123">For mail apps, the recommended image resolution is 64 x 64 pixels.</span></span> <span data-ttu-id="6f2ee-124">[HighResolutionIconUrl](highresolutioniconurl.md) 要素を使用して、高 DPI 画面で実行されている Office ホスト アプリケーションで使用するアイコンも指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="6f2ee-124">You should also specify an icon for use with Office host applications running on high DPI screens using the [HighResolutionIconUrl](highresolutioniconurl.md) element.</span></span> <span data-ttu-id="6f2ee-125">詳細については、「[効果的な AppSource と Office 内の登録リストを作成する](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity)」の「_アプリに一貫性のあるビジュアル ID を作成する_」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="6f2ee-125">For more information, see the section _Create a consistent visual identity for your app_ in [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span></span>
