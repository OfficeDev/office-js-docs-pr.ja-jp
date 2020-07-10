---
title: マニフェスト ファイルの Action 要素
description: この要素は、ユーザーがボタンまたはメニューコントロールを選択したときに実行するアクションを指定します。
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: 92c783a15d104aba0adb722ab887391b4511ebed
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094450"
---
# <a name="action-element"></a><span data-ttu-id="0f1ed-103">Action 要素</span><span class="sxs-lookup"><span data-stu-id="0f1ed-103">Action element</span></span>

<span data-ttu-id="0f1ed-104">ユーザーが[ボタン](control.md#button-control)または[メニュー](control.md#menu-dropdown-button-controls)コントロールを選択したときに実行するアクションを指定します。</span><span class="sxs-lookup"><span data-stu-id="0f1ed-104">Specifies the action to perform when the user selects a  [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) control.</span></span>

## <a name="attributes"></a><span data-ttu-id="0f1ed-105">属性</span><span class="sxs-lookup"><span data-stu-id="0f1ed-105">Attributes</span></span>

|  <span data-ttu-id="0f1ed-106">属性</span><span class="sxs-lookup"><span data-stu-id="0f1ed-106">Attribute</span></span>  |  <span data-ttu-id="0f1ed-107">必須</span><span class="sxs-lookup"><span data-stu-id="0f1ed-107">Required</span></span>  |  <span data-ttu-id="0f1ed-108">説明</span><span class="sxs-lookup"><span data-stu-id="0f1ed-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="0f1ed-109">xsi:type</span><span class="sxs-lookup"><span data-stu-id="0f1ed-109">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="0f1ed-110">はい</span><span class="sxs-lookup"><span data-stu-id="0f1ed-110">Yes</span></span>  | <span data-ttu-id="0f1ed-111">実行する操作の種類</span><span class="sxs-lookup"><span data-stu-id="0f1ed-111">Action type to take</span></span>|

## <a name="child-elements"></a><span data-ttu-id="0f1ed-112">子要素</span><span class="sxs-lookup"><span data-stu-id="0f1ed-112">Child elements</span></span>

|  <span data-ttu-id="0f1ed-113">要素</span><span class="sxs-lookup"><span data-stu-id="0f1ed-113">Element</span></span> |  <span data-ttu-id="0f1ed-114">説明</span><span class="sxs-lookup"><span data-stu-id="0f1ed-114">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="0f1ed-115">FunctionName</span><span class="sxs-lookup"><span data-stu-id="0f1ed-115">FunctionName</span></span>](#functionname) |    <span data-ttu-id="0f1ed-116">実行する関数の名前を指定します。</span><span class="sxs-lookup"><span data-stu-id="0f1ed-116">Specifies the name of the function to execute.</span></span> |
|  [<span data-ttu-id="0f1ed-117">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="0f1ed-117">SourceLocation</span></span>](#sourcelocation) |    <span data-ttu-id="0f1ed-118">この操作のソース ファイルの場所を指定します。</span><span class="sxs-lookup"><span data-stu-id="0f1ed-118">Specifies the source file location for this action.</span></span> |
| <span data-ttu-id="0f1ed-119"> [TaskpaneId](#taskpaneid)</span><span class="sxs-lookup"><span data-stu-id="0f1ed-119"> [TaskpaneId](#taskpaneid)</span></span> | <span data-ttu-id="0f1ed-120">作業ウィンドウ コンテナーの ID を指定します。</span><span class="sxs-lookup"><span data-stu-id="0f1ed-120">Specifies the ID of the task pane container.</span></span>|
| <span data-ttu-id="0f1ed-121"> [Title](#title)</span><span class="sxs-lookup"><span data-stu-id="0f1ed-121"> [Title](#title)</span></span> | <span data-ttu-id="0f1ed-122">作業ウィンドウのカスタム タイトルを指定します。</span><span class="sxs-lookup"><span data-stu-id="0f1ed-122">Specifies the custom title for the task pane.</span></span>|
| <span data-ttu-id="0f1ed-123"> [SupportsPinning](#supportspinning)</span><span class="sxs-lookup"><span data-stu-id="0f1ed-123"> [SupportsPinning](#supportspinning)</span></span> | <span data-ttu-id="0f1ed-124">作業ウィンドウがピン留めをサポートすることを指定します。これにより、ユーザーが選択を変更したときも作業ウィンドウが開いたままになります。</span><span class="sxs-lookup"><span data-stu-id="0f1ed-124">Specifies that a task pane supports pinning, which keeps the task pane open when the user changes the selection.</span></span>|
  

## <a name="xsitype"></a><span data-ttu-id="0f1ed-125">xsi:type</span><span class="sxs-lookup"><span data-stu-id="0f1ed-125">xsi:type</span></span>

<span data-ttu-id="0f1ed-126">This attribute specifies the kind of action performed when the user selects the button.</span><span class="sxs-lookup"><span data-stu-id="0f1ed-126">This attribute specifies the kind of action performed when the user selects the button.</span></span> <span data-ttu-id="0f1ed-127">It can be one of the following:</span><span class="sxs-lookup"><span data-stu-id="0f1ed-127">It can be one of the following:</span></span>

- `ExecuteFunction`
- `ShowTaskpane`

## <a name="functionname"></a><span data-ttu-id="0f1ed-128">FunctionName</span><span class="sxs-lookup"><span data-stu-id="0f1ed-128">FunctionName</span></span>

<span data-ttu-id="0f1ed-129">Required element when **xsi:type** is "ExecuteFunction".</span><span class="sxs-lookup"><span data-stu-id="0f1ed-129">Required element when **xsi:type** is "ExecuteFunction".</span></span> <span data-ttu-id="0f1ed-130">Specifies the name of the function to execute.</span><span class="sxs-lookup"><span data-stu-id="0f1ed-130">Specifies the name of the function to execute.</span></span> <span data-ttu-id="0f1ed-131">The function is contained in the file specified in the [FunctionFile](functionfile.md) element.</span><span class="sxs-lookup"><span data-stu-id="0f1ed-131">The function is contained in the file specified in the [FunctionFile](functionfile.md) element.</span></span>

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a><span data-ttu-id="0f1ed-132">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="0f1ed-132">SourceLocation</span></span>

<span data-ttu-id="0f1ed-133">**Xsi: type**が "showtaskpane" の場合に必要な要素。</span><span class="sxs-lookup"><span data-stu-id="0f1ed-133">Required element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="0f1ed-134">このアクションのソース ファイルの場所を指定します。</span><span class="sxs-lookup"><span data-stu-id="0f1ed-134">Specifies the source file location for this action.</span></span> <span data-ttu-id="0f1ed-135">**resid** 属性は、 **Resources** 要素の **Urls** 要素にある **Url** 要素の [id](resources.md) 属性の値を指定します。</span><span class="sxs-lookup"><span data-stu-id="0f1ed-135">The **resid** attribute must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the [Resources](resources.md) element.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a><span data-ttu-id="0f1ed-136">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="0f1ed-136">TaskpaneId</span></span>

<span data-ttu-id="0f1ed-137"> **xsi:type** が "ShowTaskpane" の場合に省略可能な要素。</span><span class="sxs-lookup"><span data-stu-id="0f1ed-137">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="0f1ed-138">作業ウィンドウ コンテナーの ID を指定します。</span><span class="sxs-lookup"><span data-stu-id="0f1ed-138">Specifies the ID of the task pane container.</span></span> <span data-ttu-id="0f1ed-139">複数の "ShowTaskpane" の操作があり、それぞれに対して独立したウィンドウを開く場合は、異なる **TaskpaneId** を使用します。</span><span class="sxs-lookup"><span data-stu-id="0f1ed-139">When you have multiple "ShowTaskpane" actions, use a different **TaskpaneId** if you want an independent pane for each.</span></span> <span data-ttu-id="0f1ed-140">同じウィンドウを共有する異なる操作に対しては、同じ **TaskpaneId** を使用します。</span><span class="sxs-lookup"><span data-stu-id="0f1ed-140">Use the same **TaskpaneId** for  different actions that share the same pane.</span></span> <span data-ttu-id="0f1ed-141">ユーザーが同じ **TaskpaneId** を共有するコマンドを選択した場合、ウィンドウ コンテナーは開いたままですが、ウィンドウのコンテンツは対応する Action "SourceLocation" に置き換えられます。</span><span class="sxs-lookup"><span data-stu-id="0f1ed-141">When users choose commands that share the same **TaskpaneId**, the pane container will remain open but the contents of the pane will be replaced with the corresponding Action "SourceLocation".</span></span>

> [!NOTE]
> <span data-ttu-id="0f1ed-142">この要素は、Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0f1ed-142">This element is not supported in Outlook.</span></span>

<span data-ttu-id="0f1ed-143">次の例では、同じ **TaskpaneId** を共有する 2 つのアクションを示します。</span><span class="sxs-lookup"><span data-stu-id="0f1ed-143">The following example shows two actions that share the same **TaskpaneId**.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <TaskpaneId>MyPane</TaskpaneId>
  <SourceLocation resid="aTaskPaneUrl" />
</Action>

<Action xsi:type="ShowTaskpane">
  <TaskpaneId>MyPane</TaskpaneId>
  <SourceLocation resid="anotherTaskPaneUrl" />
</Action>
```  

<span data-ttu-id="0f1ed-144">The following examples show two actions that use a different **TaskpaneId**.</span><span class="sxs-lookup"><span data-stu-id="0f1ed-144">The following examples show two actions that use a different **TaskpaneId**.</span></span> <span data-ttu-id="0f1ed-145">To see these examples in context, see [Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span><span class="sxs-lookup"><span data-stu-id="0f1ed-145">To see these examples in context, see [Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span></span>

```xml
<Action xsi:type="ShowTaskpane">
   <TaskpaneId>MyTaskPaneID1</TaskpaneId>
   <SourceLocation resid="Contoso.Taskpane1.Url" />
</Action>

<Action xsi:type="ShowTaskpane">
   <TaskpaneId>MyTaskPaneID2</TaskpaneId>
   <SourceLocation resid="Contoso.Taskpane2.Url" />
</Action>
```  

```xml
<bt:Urls>
   <bt:Url id="Contoso.Taskpane1.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane.html" />
   <bt:Url id="Contoso.Taskpane2.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane2.html" />
</bt:Urls>
```  

## <a name="title"></a><span data-ttu-id="0f1ed-146">役職</span><span class="sxs-lookup"><span data-stu-id="0f1ed-146">Title</span></span>

<span data-ttu-id="0f1ed-147"> **xsi:type** が "ShowTaskpane" の場合に省略可能な要素。</span><span class="sxs-lookup"><span data-stu-id="0f1ed-147">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="0f1ed-148">この操作に関する、作業ウィンドウのカスタム タイトルを指定します。</span><span class="sxs-lookup"><span data-stu-id="0f1ed-148">Specifies the custom title for the task pane for this action.</span></span>

<span data-ttu-id="0f1ed-149">次の例は、 **Title**要素を使用するアクションを示しています。</span><span class="sxs-lookup"><span data-stu-id="0f1ed-149">The following example shows an action that uses the **Title** element.</span></span> <span data-ttu-id="0f1ed-150">**タイトル**を文字列に直接割り当てることはないことに注意してください。</span><span class="sxs-lookup"><span data-stu-id="0f1ed-150">Note that you don't assign the **Title** to a string directly.</span></span> <span data-ttu-id="0f1ed-151">代わりに、マニフェストの [**リソース**] セクションで定義されたリソース ID (resid) を割り当てます。</span><span class="sxs-lookup"><span data-stu-id="0f1ed-151">Instead, you assign it a resource ID (resid), that is defined in the **Resources** section of the manifest.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
    <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
    <SourceLocation resid="PG.Code.Url" />
    <Title resid="PG.CodeCommand.Title" />
</Action>

 ... Other markup omitted ...
<Resources>
    <bt:Images> ...
    </bt:Images>
    <bt:Urls>
        <bt:Url id="PG.Code.Url" DefaultValue="https://localhost:3000?commands=1" />
    </bt:Urls>
    <bt:ShortStrings>
        <bt:String id="PG.CodeCommand.Title" DefaultValue="Code" />
    </bt:ShortStrings>
 ... Other markup omitted ...
</Resources>
```

## <a name="supportspinning"></a><span data-ttu-id="0f1ed-152">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="0f1ed-152">SupportsPinning</span></span>

<span data-ttu-id="0f1ed-153">**xsi:type** が "ShowTaskpane" の場合に省略可能な要素。</span><span class="sxs-lookup"><span data-stu-id="0f1ed-153">Optional element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="0f1ed-154">これを収容している [VersionOverrides](versionoverrides.md) 要素は、`xsi:type` 属性の値が `VersionOverridesV1_1` になっている必要があります。</span><span class="sxs-lookup"><span data-stu-id="0f1ed-154">The containing [VersionOverrides](versionoverrides.md) elements must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span> <span data-ttu-id="0f1ed-155">作業ウィンドウのピン留めをサポートする場合は、この要素に `true` の値を含めます。</span><span class="sxs-lookup"><span data-stu-id="0f1ed-155">Include this element with a value of `true` to support task pane pinning.</span></span> <span data-ttu-id="0f1ed-156">ユーザーは、作業ウィンドウをピン留めできるようになります。ピン留めすると、選択を変更したときも作業ウィンドウが開いたままになります。</span><span class="sxs-lookup"><span data-stu-id="0f1ed-156">The user will be able to "pin" the task pane, causing it to stay open when changing the selection.</span></span> <span data-ttu-id="0f1ed-157">詳細については、「[Outlook にピン留め可能な作業ウィンドウを実装する](../../outlook/pinnable-taskpane.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0f1ed-157">For more information, see [Implement a pinnable task pane in Outlook](../../outlook/pinnable-taskpane.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="0f1ed-158">`SupportsPinning`この要素は[要件セット 1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)で導入されましたが、現時点では、次のものを使用した Microsoft 365 サブスクライバーでのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="0f1ed-158">Although the `SupportsPinning` element was introduced in [requirement set 1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), it's currently only supported for Microsoft 365 subscribers using the following.</span></span>
> - <span data-ttu-id="0f1ed-159">Outlook 2016 以降 (ビルド7628.1000 以降)</span><span class="sxs-lookup"><span data-stu-id="0f1ed-159">Outlook 2016 or later on Windows (build 7628.1000 or later)</span></span>
> - <span data-ttu-id="0f1ed-160">Outlook 2016 以降 Mac (ビルド16.13.503 以降)</span><span class="sxs-lookup"><span data-stu-id="0f1ed-160">Outlook 2016 or later on Mac (build 16.13.503 or later)</span></span>
> - <span data-ttu-id="0f1ed-161">モダン Outlook on the web</span><span class="sxs-lookup"><span data-stu-id="0f1ed-161">Modern Outlook on the web</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```
