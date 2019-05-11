---
title: マニフェスト ファイルの OfficeTab 要素
description: ''
ms.date: 05/08/2019
localization_priority: Normal
ms.openlocfilehash: 1bf9f1d1e08a8147b52f93923229ef8fb8556fcf
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952272"
---
# <a name="officetab-element"></a><span data-ttu-id="23d2a-102">OfficeTab 要素</span><span class="sxs-lookup"><span data-stu-id="23d2a-102">OfficeTab element</span></span>

<span data-ttu-id="23d2a-p101">アドイン コマンドを表示するリボン タブを定義します。 これは既定のタブ (**[ホーム]**、**[メッセージ]**、または **[会議]** のいずれか) か、アドインで定義されたカスタム タブになります。 この要素は必須です。</span><span class="sxs-lookup"><span data-stu-id="23d2a-p101">Defines the ribbon tab on which your add-in command appears. This can either be the default tab (either  **Home**,  **Message**, or  **Meeting**), or a custom tab defined by the add-in. This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="23d2a-106">子要素</span><span class="sxs-lookup"><span data-stu-id="23d2a-106">Child elements</span></span>

|  <span data-ttu-id="23d2a-107">要素</span><span class="sxs-lookup"><span data-stu-id="23d2a-107">Element</span></span> |  <span data-ttu-id="23d2a-108">必須</span><span class="sxs-lookup"><span data-stu-id="23d2a-108">Required</span></span>  |  <span data-ttu-id="23d2a-109">説明</span><span class="sxs-lookup"><span data-stu-id="23d2a-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="23d2a-110">グループ</span><span class="sxs-lookup"><span data-stu-id="23d2a-110">Group</span></span>      | <span data-ttu-id="23d2a-111">はい</span><span class="sxs-lookup"><span data-stu-id="23d2a-111">Yes</span></span> |  <span data-ttu-id="23d2a-p102">コマンドのグループを定義します。 既定のタブには、アドインごとに 1 つのグループのみを追加できます。</span><span class="sxs-lookup"><span data-stu-id="23d2a-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="23d2a-114">ホストごとの有効なタブ `id` 値は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="23d2a-114">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="23d2a-115">**太字**の値は、デスクトップとオンラインの両方でサポートされています (たとえば、word 2016 以降の Windows および word online)。</span><span class="sxs-lookup"><span data-stu-id="23d2a-115">Values in **bold** are supported in both desktop and online (for example, Word 2016 or later on Windows and Word Online).</span></span>

### <a name="outlook"></a><span data-ttu-id="23d2a-116">Outlook</span><span class="sxs-lookup"><span data-stu-id="23d2a-116">Outlook</span></span>

- <span data-ttu-id="23d2a-117">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="23d2a-117">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="23d2a-118">Word</span><span class="sxs-lookup"><span data-stu-id="23d2a-118">Word</span></span>

- <span data-ttu-id="23d2a-119">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="23d2a-119">**TabHome**</span></span>
- <span data-ttu-id="23d2a-120">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="23d2a-120">**TabInsert**</span></span>
- <span data-ttu-id="23d2a-121">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="23d2a-121">TabWordDesign</span></span>
- <span data-ttu-id="23d2a-122">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="23d2a-122">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="23d2a-123">TabReferences</span><span class="sxs-lookup"><span data-stu-id="23d2a-123">TabReferences</span></span>
- <span data-ttu-id="23d2a-124">TabMailings</span><span class="sxs-lookup"><span data-stu-id="23d2a-124">TabMailings</span></span>
- <span data-ttu-id="23d2a-125">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="23d2a-125">TabReviewWord</span></span>
- <span data-ttu-id="23d2a-126">**TabView**</span><span class="sxs-lookup"><span data-stu-id="23d2a-126">**TabView**</span></span>
- <span data-ttu-id="23d2a-127">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="23d2a-127">TabDeveloper</span></span>
- <span data-ttu-id="23d2a-128">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="23d2a-128">TabAddIns</span></span>
- <span data-ttu-id="23d2a-129">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="23d2a-129">TabBlogPost</span></span>
- <span data-ttu-id="23d2a-130">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="23d2a-130">TabBlogInsert</span></span>
- <span data-ttu-id="23d2a-131">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="23d2a-131">TabPrintPreview</span></span>
- <span data-ttu-id="23d2a-132">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="23d2a-132">TabOutlining</span></span>
- <span data-ttu-id="23d2a-133">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="23d2a-133">TabConflicts</span></span>
- <span data-ttu-id="23d2a-134">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="23d2a-134">TabBackgroundRemoval</span></span>
- <span data-ttu-id="23d2a-135">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="23d2a-135">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="23d2a-136">Excel</span><span class="sxs-lookup"><span data-stu-id="23d2a-136">Excel</span></span>

- <span data-ttu-id="23d2a-137">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="23d2a-137">**TabHome**</span></span>
- <span data-ttu-id="23d2a-138">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="23d2a-138">**TabInsert**</span></span>
- <span data-ttu-id="23d2a-139">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="23d2a-139">TabPageLayoutExcel</span></span>
- <span data-ttu-id="23d2a-140">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="23d2a-140">TabFormulas</span></span>
- <span data-ttu-id="23d2a-141">**TabData**</span><span class="sxs-lookup"><span data-stu-id="23d2a-141">**TabData**</span></span>
- <span data-ttu-id="23d2a-142">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="23d2a-142">**TabReview**</span></span>
- <span data-ttu-id="23d2a-143">**TabView**</span><span class="sxs-lookup"><span data-stu-id="23d2a-143">**TabView**</span></span>
- <span data-ttu-id="23d2a-144">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="23d2a-144">TabDeveloper</span></span>
- <span data-ttu-id="23d2a-145">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="23d2a-145">TabAddIns</span></span>
- <span data-ttu-id="23d2a-146">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="23d2a-146">TabPrintPreview</span></span>
- <span data-ttu-id="23d2a-147">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="23d2a-147">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="23d2a-148">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="23d2a-148">PowerPoint</span></span>

- <span data-ttu-id="23d2a-149">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="23d2a-149">**TabHome**</span></span>
- <span data-ttu-id="23d2a-150">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="23d2a-150">**TabInsert**</span></span>
- <span data-ttu-id="23d2a-151">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="23d2a-151">**TabDesign**</span></span>
- <span data-ttu-id="23d2a-152">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="23d2a-152">**TabTransitions**</span></span>
- <span data-ttu-id="23d2a-153">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="23d2a-153">**TabAnimations**</span></span>
- <span data-ttu-id="23d2a-154">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="23d2a-154">TabSlideShow</span></span>
- <span data-ttu-id="23d2a-155">TabReview</span><span class="sxs-lookup"><span data-stu-id="23d2a-155">TabReview</span></span>
- <span data-ttu-id="23d2a-156">**TabView**</span><span class="sxs-lookup"><span data-stu-id="23d2a-156">**TabView**</span></span>
- <span data-ttu-id="23d2a-157">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="23d2a-157">TabDeveloper</span></span>
- <span data-ttu-id="23d2a-158">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="23d2a-158">TabAddIns</span></span>
- <span data-ttu-id="23d2a-159">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="23d2a-159">TabPrintPreview</span></span>
- <span data-ttu-id="23d2a-160">TabMerge</span><span class="sxs-lookup"><span data-stu-id="23d2a-160">TabMerge</span></span>
- <span data-ttu-id="23d2a-161">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="23d2a-161">TabGrayscale</span></span>
- <span data-ttu-id="23d2a-162">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="23d2a-162">TabBlackAndWhite</span></span>
- <span data-ttu-id="23d2a-163">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="23d2a-163">TabBroadcastPresentation</span></span>
- <span data-ttu-id="23d2a-164">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="23d2a-164">TabSlideMaster</span></span>
- <span data-ttu-id="23d2a-165">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="23d2a-165">TabHandoutMaster</span></span>
- <span data-ttu-id="23d2a-166">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="23d2a-166">TabNotesMaster</span></span>
- <span data-ttu-id="23d2a-167">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="23d2a-167">TabBackgroundRemoval</span></span>
- <span data-ttu-id="23d2a-168">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="23d2a-168">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="23d2a-169">OneNote</span><span class="sxs-lookup"><span data-stu-id="23d2a-169">OneNote</span></span>

- <span data-ttu-id="23d2a-170">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="23d2a-170">**TabHome**</span></span>
- <span data-ttu-id="23d2a-171">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="23d2a-171">**TabInsert**</span></span>
- <span data-ttu-id="23d2a-172">**TabView**</span><span class="sxs-lookup"><span data-stu-id="23d2a-172">**TabView**</span></span>
- <span data-ttu-id="23d2a-173">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="23d2a-173">TabDeveloper</span></span>
- <span data-ttu-id="23d2a-174">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="23d2a-174">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="23d2a-175">Group</span><span class="sxs-lookup"><span data-stu-id="23d2a-175">Group</span></span>

<span data-ttu-id="23d2a-p104">タブの UI 拡張ポイントのグループ。 1 つのグループに、最大 6 個のコントロールを指定できます。 **id** 属性は必須であり、各 **id** 属性はマニフェスト内で一意でなければなりません。 **id** は最大 125 文字の文字列です。 「[Group 要素](group.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="23d2a-p104">A group of UI extension points in a tab. A group can have up to six controls. The  **id** attribute is required and each **id** must be unique within the manifest. The **id** is a string with a maximum of 125 characters. See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="23d2a-180">OfficeTab の例</span><span class="sxs-lookup"><span data-stu-id="23d2a-180">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
