---
title: マニフェスト ファイルの OfficeTab 要素
description: OfficeTab 要素は、アドインコマンドが表示されるリボンタブを定義します。
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 1d1810f3d3a206f72bf9544814a3fdaaa556476e
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720492"
---
# <a name="officetab-element"></a><span data-ttu-id="f4542-103">OfficeTab 要素</span><span class="sxs-lookup"><span data-stu-id="f4542-103">OfficeTab element</span></span>

<span data-ttu-id="f4542-104">アドイン コマンドを表示するリボン タブを定義します。</span><span class="sxs-lookup"><span data-stu-id="f4542-104">Defines the ribbon tab on which your add-in command appears.</span></span> <span data-ttu-id="f4542-105">これは、既定のタブ ([**ホーム**]、[**メッセージ**]、または [**会議**]) にするか、アドインで定義されたカスタムタブにすることができます。</span><span class="sxs-lookup"><span data-stu-id="f4542-105">This can either be the default tab (either **Home**, **Message**, or **Meeting**), or a custom tab defined by the add-in.</span></span> <span data-ttu-id="f4542-106">この要素は必須です。</span><span class="sxs-lookup"><span data-stu-id="f4542-106">This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="f4542-107">子要素</span><span class="sxs-lookup"><span data-stu-id="f4542-107">Child elements</span></span>

|  <span data-ttu-id="f4542-108">要素</span><span class="sxs-lookup"><span data-stu-id="f4542-108">Element</span></span> |  <span data-ttu-id="f4542-109">必須</span><span class="sxs-lookup"><span data-stu-id="f4542-109">Required</span></span>  |  <span data-ttu-id="f4542-110">説明</span><span class="sxs-lookup"><span data-stu-id="f4542-110">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="f4542-111">グループ</span><span class="sxs-lookup"><span data-stu-id="f4542-111">Group</span></span>      | <span data-ttu-id="f4542-112">はい</span><span class="sxs-lookup"><span data-stu-id="f4542-112">Yes</span></span> |  <span data-ttu-id="f4542-p102">コマンドのグループを定義します。 既定のタブには、アドインごとに 1 つのグループのみを追加できます。</span><span class="sxs-lookup"><span data-stu-id="f4542-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="f4542-115">ホストごとの有効なタブ `id` 値は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="f4542-115">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="f4542-116">**太字**の値は、デスクトップとオンラインの両方でサポートされています (たとえば、word 2016 以降の Windows および web 上の word)。</span><span class="sxs-lookup"><span data-stu-id="f4542-116">Values in **bold** are supported in both desktop and online (for example, Word 2016 or later on Windows and Word on the web).</span></span>

### <a name="outlook"></a><span data-ttu-id="f4542-117">Outlook</span><span class="sxs-lookup"><span data-stu-id="f4542-117">Outlook</span></span>

- <span data-ttu-id="f4542-118">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="f4542-118">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="f4542-119">Word</span><span class="sxs-lookup"><span data-stu-id="f4542-119">Word</span></span>

- <span data-ttu-id="f4542-120">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="f4542-120">**TabHome**</span></span>
- <span data-ttu-id="f4542-121">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="f4542-121">**TabInsert**</span></span>
- <span data-ttu-id="f4542-122">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="f4542-122">TabWordDesign</span></span>
- <span data-ttu-id="f4542-123">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="f4542-123">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="f4542-124">TabReferences</span><span class="sxs-lookup"><span data-stu-id="f4542-124">TabReferences</span></span>
- <span data-ttu-id="f4542-125">TabMailings</span><span class="sxs-lookup"><span data-stu-id="f4542-125">TabMailings</span></span>
- <span data-ttu-id="f4542-126">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="f4542-126">TabReviewWord</span></span>
- <span data-ttu-id="f4542-127">**TabView**</span><span class="sxs-lookup"><span data-stu-id="f4542-127">**TabView**</span></span>
- <span data-ttu-id="f4542-128">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="f4542-128">TabDeveloper</span></span>
- <span data-ttu-id="f4542-129">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="f4542-129">TabAddIns</span></span>
- <span data-ttu-id="f4542-130">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="f4542-130">TabBlogPost</span></span>
- <span data-ttu-id="f4542-131">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="f4542-131">TabBlogInsert</span></span>
- <span data-ttu-id="f4542-132">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="f4542-132">TabPrintPreview</span></span>
- <span data-ttu-id="f4542-133">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="f4542-133">TabOutlining</span></span>
- <span data-ttu-id="f4542-134">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="f4542-134">TabConflicts</span></span>
- <span data-ttu-id="f4542-135">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="f4542-135">TabBackgroundRemoval</span></span>
- <span data-ttu-id="f4542-136">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="f4542-136">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="f4542-137">Excel</span><span class="sxs-lookup"><span data-stu-id="f4542-137">Excel</span></span>

- <span data-ttu-id="f4542-138">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="f4542-138">**TabHome**</span></span>
- <span data-ttu-id="f4542-139">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="f4542-139">**TabInsert**</span></span>
- <span data-ttu-id="f4542-140">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="f4542-140">TabPageLayoutExcel</span></span>
- <span data-ttu-id="f4542-141">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="f4542-141">TabFormulas</span></span>
- <span data-ttu-id="f4542-142">**TabData**</span><span class="sxs-lookup"><span data-stu-id="f4542-142">**TabData**</span></span>
- <span data-ttu-id="f4542-143">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="f4542-143">**TabReview**</span></span>
- <span data-ttu-id="f4542-144">**TabView**</span><span class="sxs-lookup"><span data-stu-id="f4542-144">**TabView**</span></span>
- <span data-ttu-id="f4542-145">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="f4542-145">TabDeveloper</span></span>
- <span data-ttu-id="f4542-146">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="f4542-146">TabAddIns</span></span>
- <span data-ttu-id="f4542-147">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="f4542-147">TabPrintPreview</span></span>
- <span data-ttu-id="f4542-148">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="f4542-148">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="f4542-149">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="f4542-149">PowerPoint</span></span>

- <span data-ttu-id="f4542-150">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="f4542-150">**TabHome**</span></span>
- <span data-ttu-id="f4542-151">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="f4542-151">**TabInsert**</span></span>
- <span data-ttu-id="f4542-152">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="f4542-152">**TabDesign**</span></span>
- <span data-ttu-id="f4542-153">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="f4542-153">**TabTransitions**</span></span>
- <span data-ttu-id="f4542-154">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="f4542-154">**TabAnimations**</span></span>
- <span data-ttu-id="f4542-155">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="f4542-155">TabSlideShow</span></span>
- <span data-ttu-id="f4542-156">TabReview</span><span class="sxs-lookup"><span data-stu-id="f4542-156">TabReview</span></span>
- <span data-ttu-id="f4542-157">**TabView**</span><span class="sxs-lookup"><span data-stu-id="f4542-157">**TabView**</span></span>
- <span data-ttu-id="f4542-158">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="f4542-158">TabDeveloper</span></span>
- <span data-ttu-id="f4542-159">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="f4542-159">TabAddIns</span></span>
- <span data-ttu-id="f4542-160">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="f4542-160">TabPrintPreview</span></span>
- <span data-ttu-id="f4542-161">TabMerge</span><span class="sxs-lookup"><span data-stu-id="f4542-161">TabMerge</span></span>
- <span data-ttu-id="f4542-162">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="f4542-162">TabGrayscale</span></span>
- <span data-ttu-id="f4542-163">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="f4542-163">TabBlackAndWhite</span></span>
- <span data-ttu-id="f4542-164">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="f4542-164">TabBroadcastPresentation</span></span>
- <span data-ttu-id="f4542-165">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="f4542-165">TabSlideMaster</span></span>
- <span data-ttu-id="f4542-166">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="f4542-166">TabHandoutMaster</span></span>
- <span data-ttu-id="f4542-167">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="f4542-167">TabNotesMaster</span></span>
- <span data-ttu-id="f4542-168">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="f4542-168">TabBackgroundRemoval</span></span>
- <span data-ttu-id="f4542-169">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="f4542-169">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="f4542-170">OneNote</span><span class="sxs-lookup"><span data-stu-id="f4542-170">OneNote</span></span>

- <span data-ttu-id="f4542-171">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="f4542-171">**TabHome**</span></span>
- <span data-ttu-id="f4542-172">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="f4542-172">**TabInsert**</span></span>
- <span data-ttu-id="f4542-173">**TabView**</span><span class="sxs-lookup"><span data-stu-id="f4542-173">**TabView**</span></span>
- <span data-ttu-id="f4542-174">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="f4542-174">TabDeveloper</span></span>
- <span data-ttu-id="f4542-175">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="f4542-175">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="f4542-176">Group</span><span class="sxs-lookup"><span data-stu-id="f4542-176">Group</span></span>

<span data-ttu-id="f4542-177">タブ内の UI 拡張ポイントのグループ。グループは最大6つのコントロールを持つことができます。</span><span class="sxs-lookup"><span data-stu-id="f4542-177">A group of UI extension points in a tab. A group can have up to six controls.</span></span> <span data-ttu-id="f4542-178">**Id**属性は必須で、各**id**はマニフェスト内で一意である必要があります。</span><span class="sxs-lookup"><span data-stu-id="f4542-178">The **id** attribute is required and each **id** must be unique within the manifest.</span></span> <span data-ttu-id="f4542-179">**Id**は、最大125文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="f4542-179">The **id** is a string with a maximum of 125 characters.</span></span> <span data-ttu-id="f4542-180">[Group 要素](group.md)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f4542-180">See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="f4542-181">OfficeTab の例</span><span class="sxs-lookup"><span data-stu-id="f4542-181">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
