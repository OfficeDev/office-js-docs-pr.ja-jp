---
title: マニフェスト ファイルの DesktopFormFactor 要素
description: デスクトップ フォーム ファクターのアドインの設定を指定します。
ms.date: 06/15/2021
localization_priority: Normal
ms.openlocfilehash: 66673d83fd8608a1ec10492d7a944b0515de61c0
ms.sourcegitcommit: 0bf0e076f705af29193abe3dba98cbfcce17b24f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/18/2021
ms.locfileid: "53007790"
---
# <a name="desktopformfactor-element"></a><span data-ttu-id="6dccc-103">DesktopFormFactor 要素</span><span class="sxs-lookup"><span data-stu-id="6dccc-103">DesktopFormFactor element</span></span>

<span data-ttu-id="6dccc-104">デスクトップ フォーム ファクターのアドインの設定を指定します。</span><span class="sxs-lookup"><span data-stu-id="6dccc-104">Specifies the settings for an add-in for the desktop form factor.</span></span> <span data-ttu-id="6dccc-105">デスクトップ フォーム ファクターには、Office on the web、Windows Mac が含まれます。</span><span class="sxs-lookup"><span data-stu-id="6dccc-105">The desktop form factor includes Office on the web, Windows, and Mac.</span></span> <span data-ttu-id="6dccc-106">Resources ノードを除く、デスクトップ フォーム ファクターのすべてのアドイン情報が **含** まれる。</span><span class="sxs-lookup"><span data-stu-id="6dccc-106">It contains all the add-in information for the desktop form factor except for the **Resources** node.</span></span>

<span data-ttu-id="6dccc-107">各 DesktopFormFactor 定義には **、FunctionFile** 要素と 1 つ以上の **ExtensionPoint 要素が含** まれています。</span><span class="sxs-lookup"><span data-stu-id="6dccc-107">Each DesktopFormFactor definition contains the **FunctionFile** element and one or more **ExtensionPoint** elements.</span></span> <span data-ttu-id="6dccc-108">詳細については、「[FunctionFile 要素](functionfile.md)」と「[ExtensionPoint 要素](extensionpoint.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6dccc-108">For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="6dccc-109">子要素</span><span class="sxs-lookup"><span data-stu-id="6dccc-109">Child elements</span></span>

| <span data-ttu-id="6dccc-110">要素</span><span class="sxs-lookup"><span data-stu-id="6dccc-110">Element</span></span>                               | <span data-ttu-id="6dccc-111">必須</span><span class="sxs-lookup"><span data-stu-id="6dccc-111">Required</span></span> | <span data-ttu-id="6dccc-112">説明</span><span class="sxs-lookup"><span data-stu-id="6dccc-112">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="6dccc-113">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="6dccc-113">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="6dccc-114">はい</span><span class="sxs-lookup"><span data-stu-id="6dccc-114">Yes</span></span>      | <span data-ttu-id="6dccc-115">アドインが機能を公開する場所を定義します。</span><span class="sxs-lookup"><span data-stu-id="6dccc-115">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="6dccc-116">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="6dccc-116">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="6dccc-117">はい</span><span class="sxs-lookup"><span data-stu-id="6dccc-117">Yes</span></span>      | <span data-ttu-id="6dccc-118">JavaScript 関数を含むファイルの URL。</span><span class="sxs-lookup"><span data-stu-id="6dccc-118">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="6dccc-119">GetStarted</span><span class="sxs-lookup"><span data-stu-id="6dccc-119">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="6dccc-120">いいえ</span><span class="sxs-lookup"><span data-stu-id="6dccc-120">No</span></span>       | <span data-ttu-id="6dccc-121">Word、Excel、またはアドインにアドインをインストールするときに表示される吹き出しをPowerPoint。</span><span class="sxs-lookup"><span data-stu-id="6dccc-121">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint.</span></span> |
| [<span data-ttu-id="6dccc-122">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="6dccc-122">SupportsSharedFolders</span></span>](supportssharedfolders.md) | <span data-ttu-id="6dccc-123">いいえ</span><span class="sxs-lookup"><span data-stu-id="6dccc-123">No</span></span> | <span data-ttu-id="6dccc-124">共有メールボックス (プレビュー Outlook共有フォルダー (つまり、代理アクセス) のシナリオで、アドインを使用できるかどうかを定義します。</span><span class="sxs-lookup"><span data-stu-id="6dccc-124">Defines whether the Outlook add-in is available in shared mailbox (now in preview) and shared folders (that is, delegate access) scenarios.</span></span> <span data-ttu-id="6dccc-125">既定では *false に* 設定されます。</span><span class="sxs-lookup"><span data-stu-id="6dccc-125">Set to *false* by default.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="6dccc-126">DesktopFormFactor の例</span><span class="sxs-lookup"><span data-stu-id="6dccc-126">DesktopFormFactor example</span></span>

```xml
...
<Hosts>
  <Host xsi:type="Presentation">
    <DesktopFormFactor>
      <FunctionFile resid="residDesktopFuncUrl" />
      <GetStarted>
        <!-- GetStarted callout -->
      </GetStarted>
      <ExtensionPoint xsi:type="PrimaryCommandSurface">
        <!-- Information on this extension point. -->
      </ExtensionPoint>
      <!-- Possibly more ExtensionPoint elements. -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
