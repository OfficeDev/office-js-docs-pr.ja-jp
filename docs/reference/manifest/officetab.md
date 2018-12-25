---
title: マニフェスト ファイルの OfficeTab 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 721064687c3c892b565a94e418815726cc0817f5
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432873"
---
# <a name="officetab-element"></a><span data-ttu-id="7459a-102">OfficeTab 要素</span><span class="sxs-lookup"><span data-stu-id="7459a-102">OfficeTab element</span></span>

<span data-ttu-id="7459a-p101">アドイン コマンドを表示するリボン タブを定義します。これは既定のタブ (**[ホーム]**、**[メッセージ]**、または **[会議]** のいずれか) か、アドインで定義されたカスタム タブになります。この要素は必須です。</span><span class="sxs-lookup"><span data-stu-id="7459a-p101">Defines the ribbon tab on which your add-in command appears. This can either be the default tab (either  **Home**,  **Message**, or  **Meeting**), or a custom tab defined by the add-in. This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="7459a-106">子要素</span><span class="sxs-lookup"><span data-stu-id="7459a-106">Child elements</span></span>

|  <span data-ttu-id="7459a-107">要素</span><span class="sxs-lookup"><span data-stu-id="7459a-107">Element</span></span> |  <span data-ttu-id="7459a-108">必須</span><span class="sxs-lookup"><span data-stu-id="7459a-108">Required</span></span>  |  <span data-ttu-id="7459a-109">Description</span><span class="sxs-lookup"><span data-stu-id="7459a-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="7459a-110">Group</span><span class="sxs-lookup"><span data-stu-id="7459a-110">Group</span></span>      | <span data-ttu-id="7459a-111">はい</span><span class="sxs-lookup"><span data-stu-id="7459a-111">Yes</span></span> |  <span data-ttu-id="7459a-p102">コマンドのグループを定義します。既定のタブには、アドインごとに 1 つのグループのみを追加できます。</span><span class="sxs-lookup"><span data-stu-id="7459a-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="7459a-114">ホストごとの有効なタブ `id` 値は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="7459a-114">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="7459a-115">**太字**の値は、デスクトップとオンラインの両方でサポートされています (たとえば、Windows 用の Word 2016 以降と Word Online)。</span><span class="sxs-lookup"><span data-stu-id="7459a-115">Values in **bold** are supported in both desktop and online (for example, Word 2016 for Windows and Word Online).</span></span>

### <a name="outlook"></a><span data-ttu-id="7459a-116">Outlook</span><span class="sxs-lookup"><span data-stu-id="7459a-116">Outlook</span></span>

- <span data-ttu-id="7459a-117">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="7459a-117">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="7459a-118">Word</span><span class="sxs-lookup"><span data-stu-id="7459a-118">Word</span></span>

- <span data-ttu-id="7459a-119">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="7459a-119">**TabHome**</span></span>
- <span data-ttu-id="7459a-120">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="7459a-120">**TabInsert**</span></span>
- <span data-ttu-id="7459a-121">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="7459a-121">TabWordDesign</span></span>
- <span data-ttu-id="7459a-122">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="7459a-122">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="7459a-123">TabReferences</span><span class="sxs-lookup"><span data-stu-id="7459a-123">TabReferences</span></span>
- <span data-ttu-id="7459a-124">TabMailings</span><span class="sxs-lookup"><span data-stu-id="7459a-124">TabMailings</span></span>
- <span data-ttu-id="7459a-125">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="7459a-125">TabReviewWord</span></span>
- <span data-ttu-id="7459a-126">**TabView**</span><span class="sxs-lookup"><span data-stu-id="7459a-126">**TabView**</span></span>
- <span data-ttu-id="7459a-127">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="7459a-127">TabDeveloper</span></span>
- <span data-ttu-id="7459a-128">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="7459a-128">TabAddIns</span></span>
- <span data-ttu-id="7459a-129">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="7459a-129">TabBlogPost</span></span>
- <span data-ttu-id="7459a-130">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="7459a-130">TabBlogInsert</span></span>
- <span data-ttu-id="7459a-131">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="7459a-131">TabPrintPreview</span></span>
- <span data-ttu-id="7459a-132">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="7459a-132">TabOutlining</span></span>
- <span data-ttu-id="7459a-133">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="7459a-133">TabConflicts</span></span>
- <span data-ttu-id="7459a-134">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="7459a-134">TabBackgroundRemoval</span></span>
- <span data-ttu-id="7459a-135">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="7459a-135">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="7459a-136">Excel</span><span class="sxs-lookup"><span data-stu-id="7459a-136">Excel</span></span>

- <span data-ttu-id="7459a-137">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="7459a-137">**TabHome**</span></span>
- <span data-ttu-id="7459a-138">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="7459a-138">**TabInsert**</span></span>
- <span data-ttu-id="7459a-139">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="7459a-139">TabPageLayoutExcel</span></span>
- <span data-ttu-id="7459a-140">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="7459a-140">TabFormulas</span></span>
- <span data-ttu-id="7459a-141">**TabData**</span><span class="sxs-lookup"><span data-stu-id="7459a-141">**TabData**</span></span>
- <span data-ttu-id="7459a-142">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="7459a-142">**TabReview**</span></span>
- <span data-ttu-id="7459a-143">**TabView**</span><span class="sxs-lookup"><span data-stu-id="7459a-143">**TabView**</span></span>
- <span data-ttu-id="7459a-144">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="7459a-144">TabDeveloper</span></span>
- <span data-ttu-id="7459a-145">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="7459a-145">TabAddIns</span></span>
- <span data-ttu-id="7459a-146">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="7459a-146">TabPrintPreview</span></span>
- <span data-ttu-id="7459a-147">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="7459a-147">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="7459a-148">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="7459a-148">PowerPoint</span></span>

- <span data-ttu-id="7459a-149">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="7459a-149">**TabHome**</span></span>
- <span data-ttu-id="7459a-150">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="7459a-150">**TabInsert**</span></span>
- <span data-ttu-id="7459a-151">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="7459a-151">**TabDesign**</span></span>
- <span data-ttu-id="7459a-152">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="7459a-152">**TabTransitions**</span></span>
- <span data-ttu-id="7459a-153">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="7459a-153">**TabAnimations**</span></span>
- <span data-ttu-id="7459a-154">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="7459a-154">TabSlideShow</span></span>
- <span data-ttu-id="7459a-155">TabReview</span><span class="sxs-lookup"><span data-stu-id="7459a-155">TabReview</span></span>
- <span data-ttu-id="7459a-156">**TabView**</span><span class="sxs-lookup"><span data-stu-id="7459a-156">**TabView**</span></span>
- <span data-ttu-id="7459a-157">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="7459a-157">TabDeveloper</span></span>
- <span data-ttu-id="7459a-158">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="7459a-158">TabAddIns</span></span>
- <span data-ttu-id="7459a-159">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="7459a-159">TabPrintPreview</span></span>
- <span data-ttu-id="7459a-160">TabMerge</span><span class="sxs-lookup"><span data-stu-id="7459a-160">TabMerge</span></span>
- <span data-ttu-id="7459a-161">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="7459a-161">TabGrayscale</span></span>
- <span data-ttu-id="7459a-162">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="7459a-162">TabBlackAndWhite</span></span>
- <span data-ttu-id="7459a-163">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="7459a-163">TabBroadcastPresentation</span></span>
- <span data-ttu-id="7459a-164">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="7459a-164">TabSlideMaster</span></span>
- <span data-ttu-id="7459a-165">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="7459a-165">TabHandoutMaster</span></span>
- <span data-ttu-id="7459a-166">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="7459a-166">TabNotesMaster</span></span>
- <span data-ttu-id="7459a-167">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="7459a-167">TabBackgroundRemoval</span></span>
- <span data-ttu-id="7459a-168">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="7459a-168">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="7459a-169">OneNote</span><span class="sxs-lookup"><span data-stu-id="7459a-169">OneNote</span></span>

- <span data-ttu-id="7459a-170">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="7459a-170">**TabHome**</span></span>
- <span data-ttu-id="7459a-171">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="7459a-171">**TabInsert**</span></span>
- <span data-ttu-id="7459a-172">**TabView**</span><span class="sxs-lookup"><span data-stu-id="7459a-172">**TabView**</span></span>
- <span data-ttu-id="7459a-173">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="7459a-173">TabDeveloper</span></span>
- <span data-ttu-id="7459a-174">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="7459a-174">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="7459a-175">Group</span><span class="sxs-lookup"><span data-stu-id="7459a-175">Group</span></span>

<span data-ttu-id="7459a-p104">タブの UI 拡張ポイントのグループ。1 つのグループに、最大 6 個のコントロールを指定できます。**id** 属性は必須であり、各 **id** 属性はマニフェスト内で一意でなければなりません。**id** は最大 125 文字の文字列です。「[Group 要素](group.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7459a-p104">A group of UI extension points in a tab. A group can have up to six controls. The  **id** attribute is required and each **id** must be unique within the manifest. The **id** is a string with a maximum of 125 characters. See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="7459a-180">OfficeTab の例</span><span class="sxs-lookup"><span data-stu-id="7459a-180">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
