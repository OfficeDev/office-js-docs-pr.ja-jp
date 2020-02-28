---
title: マニフェスト ファイルの OfficeTab 要素
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: b8458233ba93e98fe0bd8d51f5734b1fece65864
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324835"
---
# <a name="officetab-element"></a><span data-ttu-id="13da5-102">OfficeTab 要素</span><span class="sxs-lookup"><span data-stu-id="13da5-102">OfficeTab element</span></span>

<span data-ttu-id="13da5-103">アドイン コマンドを表示するリボン タブを定義します。</span><span class="sxs-lookup"><span data-stu-id="13da5-103">Defines the ribbon tab on which your add-in command appears.</span></span> <span data-ttu-id="13da5-104">これは、既定のタブ ([**ホーム**]、[**メッセージ**]、または [**会議**]) にするか、アドインで定義されたカスタムタブにすることができます。</span><span class="sxs-lookup"><span data-stu-id="13da5-104">This can either be the default tab (either **Home**, **Message**, or **Meeting**), or a custom tab defined by the add-in.</span></span> <span data-ttu-id="13da5-105">この要素は必須です。</span><span class="sxs-lookup"><span data-stu-id="13da5-105">This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="13da5-106">子要素</span><span class="sxs-lookup"><span data-stu-id="13da5-106">Child elements</span></span>

|  <span data-ttu-id="13da5-107">要素</span><span class="sxs-lookup"><span data-stu-id="13da5-107">Element</span></span> |  <span data-ttu-id="13da5-108">必須</span><span class="sxs-lookup"><span data-stu-id="13da5-108">Required</span></span>  |  <span data-ttu-id="13da5-109">説明</span><span class="sxs-lookup"><span data-stu-id="13da5-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="13da5-110">グループ</span><span class="sxs-lookup"><span data-stu-id="13da5-110">Group</span></span>      | <span data-ttu-id="13da5-111">はい</span><span class="sxs-lookup"><span data-stu-id="13da5-111">Yes</span></span> |  <span data-ttu-id="13da5-p102">コマンドのグループを定義します。 既定のタブには、アドインごとに 1 つのグループのみを追加できます。</span><span class="sxs-lookup"><span data-stu-id="13da5-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="13da5-114">ホストごとの有効なタブ `id` 値は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="13da5-114">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="13da5-115">**太字**の値は、デスクトップとオンラインの両方でサポートされています (たとえば、word 2016 以降の Windows および web 上の word)。</span><span class="sxs-lookup"><span data-stu-id="13da5-115">Values in **bold** are supported in both desktop and online (for example, Word 2016 or later on Windows and Word on the web).</span></span>

### <a name="outlook"></a><span data-ttu-id="13da5-116">Outlook</span><span class="sxs-lookup"><span data-stu-id="13da5-116">Outlook</span></span>

- <span data-ttu-id="13da5-117">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="13da5-117">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="13da5-118">Word</span><span class="sxs-lookup"><span data-stu-id="13da5-118">Word</span></span>

- <span data-ttu-id="13da5-119">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="13da5-119">**TabHome**</span></span>
- <span data-ttu-id="13da5-120">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="13da5-120">**TabInsert**</span></span>
- <span data-ttu-id="13da5-121">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="13da5-121">TabWordDesign</span></span>
- <span data-ttu-id="13da5-122">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="13da5-122">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="13da5-123">TabReferences</span><span class="sxs-lookup"><span data-stu-id="13da5-123">TabReferences</span></span>
- <span data-ttu-id="13da5-124">TabMailings</span><span class="sxs-lookup"><span data-stu-id="13da5-124">TabMailings</span></span>
- <span data-ttu-id="13da5-125">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="13da5-125">TabReviewWord</span></span>
- <span data-ttu-id="13da5-126">**TabView**</span><span class="sxs-lookup"><span data-stu-id="13da5-126">**TabView**</span></span>
- <span data-ttu-id="13da5-127">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="13da5-127">TabDeveloper</span></span>
- <span data-ttu-id="13da5-128">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="13da5-128">TabAddIns</span></span>
- <span data-ttu-id="13da5-129">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="13da5-129">TabBlogPost</span></span>
- <span data-ttu-id="13da5-130">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="13da5-130">TabBlogInsert</span></span>
- <span data-ttu-id="13da5-131">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="13da5-131">TabPrintPreview</span></span>
- <span data-ttu-id="13da5-132">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="13da5-132">TabOutlining</span></span>
- <span data-ttu-id="13da5-133">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="13da5-133">TabConflicts</span></span>
- <span data-ttu-id="13da5-134">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="13da5-134">TabBackgroundRemoval</span></span>
- <span data-ttu-id="13da5-135">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="13da5-135">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="13da5-136">Excel</span><span class="sxs-lookup"><span data-stu-id="13da5-136">Excel</span></span>

- <span data-ttu-id="13da5-137">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="13da5-137">**TabHome**</span></span>
- <span data-ttu-id="13da5-138">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="13da5-138">**TabInsert**</span></span>
- <span data-ttu-id="13da5-139">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="13da5-139">TabPageLayoutExcel</span></span>
- <span data-ttu-id="13da5-140">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="13da5-140">TabFormulas</span></span>
- <span data-ttu-id="13da5-141">**TabData**</span><span class="sxs-lookup"><span data-stu-id="13da5-141">**TabData**</span></span>
- <span data-ttu-id="13da5-142">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="13da5-142">**TabReview**</span></span>
- <span data-ttu-id="13da5-143">**TabView**</span><span class="sxs-lookup"><span data-stu-id="13da5-143">**TabView**</span></span>
- <span data-ttu-id="13da5-144">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="13da5-144">TabDeveloper</span></span>
- <span data-ttu-id="13da5-145">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="13da5-145">TabAddIns</span></span>
- <span data-ttu-id="13da5-146">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="13da5-146">TabPrintPreview</span></span>
- <span data-ttu-id="13da5-147">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="13da5-147">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="13da5-148">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="13da5-148">PowerPoint</span></span>

- <span data-ttu-id="13da5-149">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="13da5-149">**TabHome**</span></span>
- <span data-ttu-id="13da5-150">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="13da5-150">**TabInsert**</span></span>
- <span data-ttu-id="13da5-151">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="13da5-151">**TabDesign**</span></span>
- <span data-ttu-id="13da5-152">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="13da5-152">**TabTransitions**</span></span>
- <span data-ttu-id="13da5-153">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="13da5-153">**TabAnimations**</span></span>
- <span data-ttu-id="13da5-154">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="13da5-154">TabSlideShow</span></span>
- <span data-ttu-id="13da5-155">TabReview</span><span class="sxs-lookup"><span data-stu-id="13da5-155">TabReview</span></span>
- <span data-ttu-id="13da5-156">**TabView**</span><span class="sxs-lookup"><span data-stu-id="13da5-156">**TabView**</span></span>
- <span data-ttu-id="13da5-157">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="13da5-157">TabDeveloper</span></span>
- <span data-ttu-id="13da5-158">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="13da5-158">TabAddIns</span></span>
- <span data-ttu-id="13da5-159">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="13da5-159">TabPrintPreview</span></span>
- <span data-ttu-id="13da5-160">TabMerge</span><span class="sxs-lookup"><span data-stu-id="13da5-160">TabMerge</span></span>
- <span data-ttu-id="13da5-161">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="13da5-161">TabGrayscale</span></span>
- <span data-ttu-id="13da5-162">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="13da5-162">TabBlackAndWhite</span></span>
- <span data-ttu-id="13da5-163">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="13da5-163">TabBroadcastPresentation</span></span>
- <span data-ttu-id="13da5-164">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="13da5-164">TabSlideMaster</span></span>
- <span data-ttu-id="13da5-165">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="13da5-165">TabHandoutMaster</span></span>
- <span data-ttu-id="13da5-166">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="13da5-166">TabNotesMaster</span></span>
- <span data-ttu-id="13da5-167">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="13da5-167">TabBackgroundRemoval</span></span>
- <span data-ttu-id="13da5-168">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="13da5-168">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="13da5-169">OneNote</span><span class="sxs-lookup"><span data-stu-id="13da5-169">OneNote</span></span>

- <span data-ttu-id="13da5-170">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="13da5-170">**TabHome**</span></span>
- <span data-ttu-id="13da5-171">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="13da5-171">**TabInsert**</span></span>
- <span data-ttu-id="13da5-172">**TabView**</span><span class="sxs-lookup"><span data-stu-id="13da5-172">**TabView**</span></span>
- <span data-ttu-id="13da5-173">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="13da5-173">TabDeveloper</span></span>
- <span data-ttu-id="13da5-174">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="13da5-174">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="13da5-175">Group</span><span class="sxs-lookup"><span data-stu-id="13da5-175">Group</span></span>

<span data-ttu-id="13da5-176">タブ内の UI 拡張ポイントのグループ。グループは最大6つのコントロールを持つことができます。</span><span class="sxs-lookup"><span data-stu-id="13da5-176">A group of UI extension points in a tab. A group can have up to six controls.</span></span> <span data-ttu-id="13da5-177">**Id**属性は必須で、各**id**はマニフェスト内で一意である必要があります。</span><span class="sxs-lookup"><span data-stu-id="13da5-177">The **id** attribute is required and each **id** must be unique within the manifest.</span></span> <span data-ttu-id="13da5-178">**Id**は、最大125文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="13da5-178">The **id** is a string with a maximum of 125 characters.</span></span> <span data-ttu-id="13da5-179">[Group 要素](group.md)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="13da5-179">See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="13da5-180">OfficeTab の例</span><span class="sxs-lookup"><span data-stu-id="13da5-180">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
