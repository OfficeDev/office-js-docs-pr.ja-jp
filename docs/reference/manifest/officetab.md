---
title: マニフェスト ファイルの OfficeTab 要素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: b61c245c000f8bf13eb71c991ec57a125993c2fc
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450465"
---
# <a name="officetab-element"></a><span data-ttu-id="0a76d-102">OfficeTab 要素</span><span class="sxs-lookup"><span data-stu-id="0a76d-102">OfficeTab element</span></span>

<span data-ttu-id="0a76d-p101">アドイン コマンドを表示するリボン タブを定義します。 これは既定のタブ (**[ホーム]**、**[メッセージ]**、または **[会議]** のいずれか) か、アドインで定義されたカスタム タブになります。 この要素は必須です。</span><span class="sxs-lookup"><span data-stu-id="0a76d-p101">Defines the ribbon tab on which your add-in command appears. This can either be the default tab (either  **Home**,  **Message**, or  **Meeting**), or a custom tab defined by the add-in. This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="0a76d-106">子要素</span><span class="sxs-lookup"><span data-stu-id="0a76d-106">Child elements</span></span>

|  <span data-ttu-id="0a76d-107">要素</span><span class="sxs-lookup"><span data-stu-id="0a76d-107">Element</span></span> |  <span data-ttu-id="0a76d-108">必須</span><span class="sxs-lookup"><span data-stu-id="0a76d-108">Required</span></span>  |  <span data-ttu-id="0a76d-109">説明</span><span class="sxs-lookup"><span data-stu-id="0a76d-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="0a76d-110">グループ</span><span class="sxs-lookup"><span data-stu-id="0a76d-110">Group</span></span>      | <span data-ttu-id="0a76d-111">はい</span><span class="sxs-lookup"><span data-stu-id="0a76d-111">Yes</span></span> |  <span data-ttu-id="0a76d-p102">コマンドのグループを定義します。既定のタブには、アドインごとに 1 つのグループのみを追加できます。</span><span class="sxs-lookup"><span data-stu-id="0a76d-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="0a76d-114">ホストごとの有効なタブ `id` 値は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="0a76d-114">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="0a76d-115">**太字**の値は、デスクトップとオンラインの両方でサポートされています (たとえば、Windows 用の Word 2016 以降と Word Online)。</span><span class="sxs-lookup"><span data-stu-id="0a76d-115">Values in **bold** are supported in both desktop and online (for example, Word 2016 or later for Windows and Word Online).</span></span>

### <a name="outlook"></a><span data-ttu-id="0a76d-116">Outlook</span><span class="sxs-lookup"><span data-stu-id="0a76d-116">Outlook</span></span>

- <span data-ttu-id="0a76d-117">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="0a76d-117">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="0a76d-118">Word</span><span class="sxs-lookup"><span data-stu-id="0a76d-118">Word</span></span>

- <span data-ttu-id="0a76d-119">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="0a76d-119">**TabHome**</span></span>
- <span data-ttu-id="0a76d-120">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="0a76d-120">**TabInsert**</span></span>
- <span data-ttu-id="0a76d-121">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="0a76d-121">TabWordDesign</span></span>
- <span data-ttu-id="0a76d-122">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="0a76d-122">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="0a76d-123">TabReferences</span><span class="sxs-lookup"><span data-stu-id="0a76d-123">TabReferences</span></span>
- <span data-ttu-id="0a76d-124">TabMailings</span><span class="sxs-lookup"><span data-stu-id="0a76d-124">TabMailings</span></span>
- <span data-ttu-id="0a76d-125">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="0a76d-125">TabReviewWord</span></span>
- <span data-ttu-id="0a76d-126">**TabView**</span><span class="sxs-lookup"><span data-stu-id="0a76d-126">**TabView**</span></span>
- <span data-ttu-id="0a76d-127">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="0a76d-127">TabDeveloper</span></span>
- <span data-ttu-id="0a76d-128">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="0a76d-128">TabAddIns</span></span>
- <span data-ttu-id="0a76d-129">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="0a76d-129">TabBlogPost</span></span>
- <span data-ttu-id="0a76d-130">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="0a76d-130">TabBlogInsert</span></span>
- <span data-ttu-id="0a76d-131">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="0a76d-131">TabPrintPreview</span></span>
- <span data-ttu-id="0a76d-132">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="0a76d-132">TabOutlining</span></span>
- <span data-ttu-id="0a76d-133">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="0a76d-133">TabConflicts</span></span>
- <span data-ttu-id="0a76d-134">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="0a76d-134">TabBackgroundRemoval</span></span>
- <span data-ttu-id="0a76d-135">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="0a76d-135">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="0a76d-136">Excel</span><span class="sxs-lookup"><span data-stu-id="0a76d-136">Excel</span></span>

- <span data-ttu-id="0a76d-137">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="0a76d-137">**TabHome**</span></span>
- <span data-ttu-id="0a76d-138">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="0a76d-138">**TabInsert**</span></span>
- <span data-ttu-id="0a76d-139">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="0a76d-139">TabPageLayoutExcel</span></span>
- <span data-ttu-id="0a76d-140">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="0a76d-140">TabFormulas</span></span>
- <span data-ttu-id="0a76d-141">**TabData**</span><span class="sxs-lookup"><span data-stu-id="0a76d-141">**TabData**</span></span>
- <span data-ttu-id="0a76d-142">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="0a76d-142">**TabReview**</span></span>
- <span data-ttu-id="0a76d-143">**TabView**</span><span class="sxs-lookup"><span data-stu-id="0a76d-143">**TabView**</span></span>
- <span data-ttu-id="0a76d-144">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="0a76d-144">TabDeveloper</span></span>
- <span data-ttu-id="0a76d-145">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="0a76d-145">TabAddIns</span></span>
- <span data-ttu-id="0a76d-146">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="0a76d-146">TabPrintPreview</span></span>
- <span data-ttu-id="0a76d-147">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="0a76d-147">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="0a76d-148">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="0a76d-148">PowerPoint</span></span>

- <span data-ttu-id="0a76d-149">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="0a76d-149">**TabHome**</span></span>
- <span data-ttu-id="0a76d-150">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="0a76d-150">**TabInsert**</span></span>
- <span data-ttu-id="0a76d-151">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="0a76d-151">**TabDesign**</span></span>
- <span data-ttu-id="0a76d-152">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="0a76d-152">**TabTransitions**</span></span>
- <span data-ttu-id="0a76d-153">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="0a76d-153">**TabAnimations**</span></span>
- <span data-ttu-id="0a76d-154">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="0a76d-154">TabSlideShow</span></span>
- <span data-ttu-id="0a76d-155">TabReview</span><span class="sxs-lookup"><span data-stu-id="0a76d-155">TabReview</span></span>
- <span data-ttu-id="0a76d-156">**TabView**</span><span class="sxs-lookup"><span data-stu-id="0a76d-156">**TabView**</span></span>
- <span data-ttu-id="0a76d-157">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="0a76d-157">TabDeveloper</span></span>
- <span data-ttu-id="0a76d-158">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="0a76d-158">TabAddIns</span></span>
- <span data-ttu-id="0a76d-159">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="0a76d-159">TabPrintPreview</span></span>
- <span data-ttu-id="0a76d-160">TabMerge</span><span class="sxs-lookup"><span data-stu-id="0a76d-160">TabMerge</span></span>
- <span data-ttu-id="0a76d-161">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="0a76d-161">TabGrayscale</span></span>
- <span data-ttu-id="0a76d-162">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="0a76d-162">TabBlackAndWhite</span></span>
- <span data-ttu-id="0a76d-163">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="0a76d-163">TabBroadcastPresentation</span></span>
- <span data-ttu-id="0a76d-164">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="0a76d-164">TabSlideMaster</span></span>
- <span data-ttu-id="0a76d-165">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="0a76d-165">TabHandoutMaster</span></span>
- <span data-ttu-id="0a76d-166">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="0a76d-166">TabNotesMaster</span></span>
- <span data-ttu-id="0a76d-167">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="0a76d-167">TabBackgroundRemoval</span></span>
- <span data-ttu-id="0a76d-168">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="0a76d-168">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="0a76d-169">OneNote</span><span class="sxs-lookup"><span data-stu-id="0a76d-169">OneNote</span></span>

- <span data-ttu-id="0a76d-170">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="0a76d-170">**TabHome**</span></span>
- <span data-ttu-id="0a76d-171">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="0a76d-171">**TabInsert**</span></span>
- <span data-ttu-id="0a76d-172">**TabView**</span><span class="sxs-lookup"><span data-stu-id="0a76d-172">**TabView**</span></span>
- <span data-ttu-id="0a76d-173">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="0a76d-173">TabDeveloper</span></span>
- <span data-ttu-id="0a76d-174">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="0a76d-174">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="0a76d-175">Group</span><span class="sxs-lookup"><span data-stu-id="0a76d-175">Group</span></span>

<span data-ttu-id="0a76d-p104">タブの UI 拡張ポイントのグループ。 1 つのグループに、最大 6 個のコントロールを指定できます。 **id** 属性は必須であり、各 **id** 属性はマニフェスト内で一意でなければなりません。 **id** は最大 125 文字の文字列です。 「[Group 要素](group.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0a76d-p104">A group of UI extension points in a tab. A group can have up to six controls. The  **id** attribute is required and each **id** must be unique within the manifest. The **id** is a string with a maximum of 125 characters. See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="0a76d-180">OfficeTab の例</span><span class="sxs-lookup"><span data-stu-id="0a76d-180">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
