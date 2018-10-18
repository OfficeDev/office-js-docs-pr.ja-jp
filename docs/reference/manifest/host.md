# <a name="host-element"></a><span data-ttu-id="48a25-101">Host 要素</span><span class="sxs-lookup"><span data-stu-id="48a25-101">Host element</span></span>

<span data-ttu-id="48a25-102">アドインでアクティブ化する Office アプリケーションの種類を個別に指定します。</span><span class="sxs-lookup"><span data-stu-id="48a25-102">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="48a25-103">重要:**Host** 要素の構文は、要素が[基本のマニフェスト](#basic-manifest)で定義されているか、[VersionOverrides](#versionoverrides-node) ノードで定義されているかによって異なります。</span><span class="sxs-lookup"><span data-stu-id="48a25-103">Important: The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="48a25-104">ただし、機能は変わりません。</span><span class="sxs-lookup"><span data-stu-id="48a25-104">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="48a25-105">基本のマニフェスト</span><span class="sxs-lookup"><span data-stu-id="48a25-105">Basic manifest</span></span>

<span data-ttu-id="48a25-106">基本のマニフェストで定義されている場合 ([OfficeApp](officeapp.md) の下)、ホストの種類は `Name` 属性によって決定されます。</span><span class="sxs-lookup"><span data-stu-id="48a25-106">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>   

### <a name="attributes"></a><span data-ttu-id="48a25-107">属性</span><span class="sxs-lookup"><span data-stu-id="48a25-107">Attributes</span></span>

| <span data-ttu-id="48a25-108">属性</span><span class="sxs-lookup"><span data-stu-id="48a25-108">Attribute</span></span>     | <span data-ttu-id="48a25-109">型</span><span class="sxs-lookup"><span data-stu-id="48a25-109">Type</span></span>   | <span data-ttu-id="48a25-110">必須</span><span class="sxs-lookup"><span data-stu-id="48a25-110">Required</span></span> | <span data-ttu-id="48a25-111">説明</span><span class="sxs-lookup"><span data-stu-id="48a25-111">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="48a25-112">名前</span><span class="sxs-lookup"><span data-stu-id="48a25-112">Name</span></span>](#name) | <span data-ttu-id="48a25-113">文字列</span><span class="sxs-lookup"><span data-stu-id="48a25-113">string</span></span> | <span data-ttu-id="48a25-114">必須</span><span class="sxs-lookup"><span data-stu-id="48a25-114">required</span></span> | <span data-ttu-id="48a25-115">Office ホスト アプリケーションの種類の名前。</span><span class="sxs-lookup"><span data-stu-id="48a25-115">The name of the type of Office host application.</span></span> |

### <a name="name"></a><span data-ttu-id="48a25-116">名前</span><span class="sxs-lookup"><span data-stu-id="48a25-116">Name</span></span>
<span data-ttu-id="48a25-p102">このアドインが対象にするホストの種類を指定します。この値は、次のいずれかである必要があります。</span><span class="sxs-lookup"><span data-stu-id="48a25-p102">Specifies the Host type targeted by this add-in. The value must be one of the following:</span></span>

- <span data-ttu-id="48a25-119">`Document` Word</span><span class="sxs-lookup"><span data-stu-id="48a25-119">`Document` (Word)</span></span>
- <span data-ttu-id="48a25-120">`Database` Access</span><span class="sxs-lookup"><span data-stu-id="48a25-120">`Database` (Access)</span></span>
- <span data-ttu-id="48a25-121">`Mailbox` Outlook</span><span class="sxs-lookup"><span data-stu-id="48a25-121">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="48a25-122">`Notebook` OneNote</span><span class="sxs-lookup"><span data-stu-id="48a25-122">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="48a25-123">`Presentation` PowerPoint</span><span class="sxs-lookup"><span data-stu-id="48a25-123">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="48a25-124">`Project` プロジェクト</span><span class="sxs-lookup"><span data-stu-id="48a25-124">`Project` (Project)</span></span>
- <span data-ttu-id="48a25-125">`Workbook` Excel</span><span class="sxs-lookup"><span data-stu-id="48a25-125">`Workbook` (Excel)</span></span>

### <a name="example"></a><span data-ttu-id="48a25-126">例</span><span class="sxs-lookup"><span data-stu-id="48a25-126">Example</span></span>
```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="48a25-127">VersionOverrides ノード</span><span class="sxs-lookup"><span data-stu-id="48a25-127">VersionOverrides node</span></span>
<span data-ttu-id="48a25-128">[VersionOverrides](versionoverrides.md) で定義されている場合、ホストの種類は `xsi:type` 属性によって決定されます。</span><span class="sxs-lookup"><span data-stu-id="48a25-128">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span> 

### <a name="attributes"></a><span data-ttu-id="48a25-129">属性</span><span class="sxs-lookup"><span data-stu-id="48a25-129">Attributes</span></span>

|  <span data-ttu-id="48a25-130">属性</span><span class="sxs-lookup"><span data-stu-id="48a25-130">Attribute</span></span>  |  <span data-ttu-id="48a25-131">必須</span><span class="sxs-lookup"><span data-stu-id="48a25-131">Required</span></span>  |  <span data-ttu-id="48a25-132">説明</span><span class="sxs-lookup"><span data-stu-id="48a25-132">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="48a25-133">xsi:type</span><span class="sxs-lookup"><span data-stu-id="48a25-133">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="48a25-134">はい</span><span class="sxs-lookup"><span data-stu-id="48a25-134">Yes</span></span>  | <span data-ttu-id="48a25-135">これらの設定を適用する Office ホストについて説明します。</span><span class="sxs-lookup"><span data-stu-id="48a25-135">Describes the Office host where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="48a25-136">子要素</span><span class="sxs-lookup"><span data-stu-id="48a25-136">Child elements</span></span>

|  <span data-ttu-id="48a25-137">要素</span><span class="sxs-lookup"><span data-stu-id="48a25-137">Element</span></span> |  <span data-ttu-id="48a25-138">必須</span><span class="sxs-lookup"><span data-stu-id="48a25-138">Required</span></span>  |  <span data-ttu-id="48a25-139">説明</span><span class="sxs-lookup"><span data-stu-id="48a25-139">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="48a25-140">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="48a25-140">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="48a25-141">はい</span><span class="sxs-lookup"><span data-stu-id="48a25-141">Yes</span></span>   |  <span data-ttu-id="48a25-142">デスクトップ フォーム ファクターの設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="48a25-142">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="48a25-143">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="48a25-143">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="48a25-144">いいえ</span><span class="sxs-lookup"><span data-stu-id="48a25-144">No</span></span>   |  <span data-ttu-id="48a25-p103">モバイル フォーム ファクターの設定を定義します。**注:** この要素は、Outlook for iOS でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="48a25-p103">Defines the settings for the mobile form factor. **Note:** this element is only supported in Outlook for iOS.</span></span> |
|  [<span data-ttu-id="48a25-147">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="48a25-147">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="48a25-148">いいえ</span><span class="sxs-lookup"><span data-stu-id="48a25-148">No</span></span>   |  <span data-ttu-id="48a25-149">すべてのフォーム ファクターの設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="48a25-149">Defines the settings for all form factors.</span></span> <span data-ttu-id="48a25-150">Excel のカスタム関数でのみ使用します。</span><span class="sxs-lookup"><span data-stu-id="48a25-150">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="48a25-151">xsi:type</span><span class="sxs-lookup"><span data-stu-id="48a25-151">xsi:type</span></span>

<span data-ttu-id="48a25-152">含まれている設定を適用する Office ホスト (Word、Excel、PowerPoint、Outlook、OneNote) を制御します。</span><span class="sxs-lookup"><span data-stu-id="48a25-152">Controls which Office host (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="48a25-153">この値は、次のいずれかである必要があります。</span><span class="sxs-lookup"><span data-stu-id="48a25-153">The value must be one of the following:</span></span>

- <span data-ttu-id="48a25-154">`Document` Word</span><span class="sxs-lookup"><span data-stu-id="48a25-154">`Document` (Word)</span></span>
- <span data-ttu-id="48a25-155">`MailHost` Outlook</span><span class="sxs-lookup"><span data-stu-id="48a25-155">`MailHost` (Outlook)</span></span>    
- <span data-ttu-id="48a25-156">`Notebook` OneNote</span><span class="sxs-lookup"><span data-stu-id="48a25-156">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="48a25-157">`Presentation` PowerPoint</span><span class="sxs-lookup"><span data-stu-id="48a25-157">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="48a25-158">`Workbook` Excel</span><span class="sxs-lookup"><span data-stu-id="48a25-158">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="48a25-159">ホストの例</span><span class="sxs-lookup"><span data-stu-id="48a25-159">Host example</span></span> 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
