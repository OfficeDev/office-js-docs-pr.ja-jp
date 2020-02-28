---
title: マニフェスト ファイルの Action 要素
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: b05da08f4995c7d8f7270e7fba6f416c9903b066
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324892"
---
# <a name="action-element"></a><span data-ttu-id="7e02d-102">Action 要素</span><span class="sxs-lookup"><span data-stu-id="7e02d-102">Action element</span></span>

<span data-ttu-id="7e02d-103">ユーザーが[ボタン](control.md#button-control)または[メニュー](control.md#menu-dropdown-button-controls) コントロールを選択したときに実行する操作を指定します。</span><span class="sxs-lookup"><span data-stu-id="7e02d-103">Specifies the action to perform when the user selects a  [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) controls.</span></span>

## <a name="attributes"></a><span data-ttu-id="7e02d-104">属性</span><span class="sxs-lookup"><span data-stu-id="7e02d-104">Attributes</span></span>

|  <span data-ttu-id="7e02d-105">属性</span><span class="sxs-lookup"><span data-stu-id="7e02d-105">Attribute</span></span>  |  <span data-ttu-id="7e02d-106">必須</span><span class="sxs-lookup"><span data-stu-id="7e02d-106">Required</span></span>  |  <span data-ttu-id="7e02d-107">説明</span><span class="sxs-lookup"><span data-stu-id="7e02d-107">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="7e02d-108">xsi:type</span><span class="sxs-lookup"><span data-stu-id="7e02d-108">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="7e02d-109">はい</span><span class="sxs-lookup"><span data-stu-id="7e02d-109">Yes</span></span>  | <span data-ttu-id="7e02d-110">実行する操作の種類</span><span class="sxs-lookup"><span data-stu-id="7e02d-110">Action type to take</span></span>|

## <a name="child-elements"></a><span data-ttu-id="7e02d-111">子要素</span><span class="sxs-lookup"><span data-stu-id="7e02d-111">Child elements</span></span>

|  <span data-ttu-id="7e02d-112">要素</span><span class="sxs-lookup"><span data-stu-id="7e02d-112">Element</span></span> |  <span data-ttu-id="7e02d-113">説明</span><span class="sxs-lookup"><span data-stu-id="7e02d-113">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="7e02d-114">FunctionName</span><span class="sxs-lookup"><span data-stu-id="7e02d-114">FunctionName</span></span>](#functionname) |    <span data-ttu-id="7e02d-115">実行する関数の名前を指定します。</span><span class="sxs-lookup"><span data-stu-id="7e02d-115">Specifies the name of the function to execute.</span></span> |
|  [<span data-ttu-id="7e02d-116">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="7e02d-116">SourceLocation</span></span>](#sourcelocation) |    <span data-ttu-id="7e02d-117">この操作のソース ファイルの場所を指定します。</span><span class="sxs-lookup"><span data-stu-id="7e02d-117">Specifies the source file location for this action.</span></span> |
| <span data-ttu-id="7e02d-118"> [TaskpaneId](#taskpaneid)</span><span class="sxs-lookup"><span data-stu-id="7e02d-118"> [TaskpaneId](#taskpaneid)</span></span> | <span data-ttu-id="7e02d-119">作業ウィンドウ コンテナーの ID を指定します。</span><span class="sxs-lookup"><span data-stu-id="7e02d-119">Specifies the ID of the task pane container.</span></span>|
| <span data-ttu-id="7e02d-120"> [Title](#title)</span><span class="sxs-lookup"><span data-stu-id="7e02d-120"> [Title](#title)</span></span> | <span data-ttu-id="7e02d-121">作業ウィンドウのカスタム タイトルを指定します。</span><span class="sxs-lookup"><span data-stu-id="7e02d-121">Specifies the custom title for the task pane.</span></span>|
| <span data-ttu-id="7e02d-122"> [SupportsPinning](#supportspinning)</span><span class="sxs-lookup"><span data-stu-id="7e02d-122"> [SupportsPinning](#supportspinning)</span></span> | <span data-ttu-id="7e02d-123">作業ウィンドウがピン留めをサポートすることを指定します。これにより、ユーザーが選択を変更したときも作業ウィンドウが開いたままになります。</span><span class="sxs-lookup"><span data-stu-id="7e02d-123">Specifies that a task pane supports pinning, which keeps the task pane open when the user changes the selection.</span></span>|
  

## <a name="xsitype"></a><span data-ttu-id="7e02d-124">xsi:type</span><span class="sxs-lookup"><span data-stu-id="7e02d-124">xsi:type</span></span>

<span data-ttu-id="7e02d-p101">この属性は、ユーザーがボタンをクリックしたときに実行される操作の種類を指定します。次のいずれかを指定できます。</span><span class="sxs-lookup"><span data-stu-id="7e02d-p101">This attribute specifies the kind of action performed when the user selects the button. It can be one of the following:</span></span>

- `ExecuteFunction`
- `ShowTaskpane`

## <a name="functionname"></a><span data-ttu-id="7e02d-127">FunctionName</span><span class="sxs-lookup"><span data-stu-id="7e02d-127">FunctionName</span></span>

<span data-ttu-id="7e02d-p102">**xsi:type** が "ExecuteFunction" のときに必ず指定する要素です。実行する関数の名前を指定します。関数は、[FunctionFile](functionfile.md) 要素に指定されたファイルに含まれています。</span><span class="sxs-lookup"><span data-stu-id="7e02d-p102">Required element when **xsi:type** is "ExecuteFunction". Specifies the name of the function to execute. The function is contained in the file specified in the [FunctionFile](functionfile.md) element.</span></span>

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a><span data-ttu-id="7e02d-131">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="7e02d-131">SourceLocation</span></span>

<span data-ttu-id="7e02d-132">**Xsi: type**が "showtaskpane" の場合に必要な要素。</span><span class="sxs-lookup"><span data-stu-id="7e02d-132">Required element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="7e02d-133">このアクションのソース ファイルの場所を指定します。</span><span class="sxs-lookup"><span data-stu-id="7e02d-133">Specifies the source file location for this action.</span></span> <span data-ttu-id="7e02d-134">**resid** 属性は、 **Resources** 要素の **Urls** 要素にある **Url** 要素の [id](resources.md) 属性の値を指定します。</span><span class="sxs-lookup"><span data-stu-id="7e02d-134">The **resid** attribute must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the [Resources](resources.md) element.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a><span data-ttu-id="7e02d-135">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="7e02d-135">TaskpaneId</span></span>

<span data-ttu-id="7e02d-136"> **xsi:type** が "ShowTaskpane" の場合に省略可能な要素。</span><span class="sxs-lookup"><span data-stu-id="7e02d-136">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="7e02d-137">作業ウィンドウ コンテナーの ID を指定します。</span><span class="sxs-lookup"><span data-stu-id="7e02d-137">Specifies the ID of the task pane container.</span></span> <span data-ttu-id="7e02d-138">複数の "ShowTaskpane" の操作があり、それぞれに対して独立したウィンドウを開く場合は、異なる **TaskpaneId** を使用します。</span><span class="sxs-lookup"><span data-stu-id="7e02d-138">When you have multiple "ShowTaskpane" actions, use a different **TaskpaneId** if you want an independent pane for each.</span></span> <span data-ttu-id="7e02d-139">同じウィンドウを共有する異なる操作に対しては、同じ **TaskpaneId** を使用します。</span><span class="sxs-lookup"><span data-stu-id="7e02d-139">Use the same **TaskpaneId** for  different actions that share the same pane.</span></span> <span data-ttu-id="7e02d-140">ユーザーが同じ **TaskpaneId** を共有するコマンドを選択した場合、ウィンドウ コンテナーは開いたままですが、ウィンドウのコンテンツは対応する Action "SourceLocation" に置き換えられます。</span><span class="sxs-lookup"><span data-stu-id="7e02d-140">When users choose commands that share the same **TaskpaneId**, the pane container will remain open but the contents of the pane will be replaced with the corresponding Action "SourceLocation".</span></span>

> [!NOTE]
> <span data-ttu-id="7e02d-141">この要素は、Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="7e02d-141">This element is not supported in Outlook.</span></span>

<span data-ttu-id="7e02d-142">次の例では、同じ **TaskpaneId** を共有する 2 つのアクションを示します。</span><span class="sxs-lookup"><span data-stu-id="7e02d-142">The following example shows two actions that share the same **TaskpaneId**.</span></span>

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

<span data-ttu-id="7e02d-p105">次の例では、異なる **TaskpaneId** を使用する 2 つの操作を示します。これらの例を全体的な流れで確認する場合は、「[Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7e02d-p105">The following examples show two actions that use a different **TaskpaneId**. To see these examples in context, see [Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span></span>

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

## <a name="title"></a><span data-ttu-id="7e02d-145">役職</span><span class="sxs-lookup"><span data-stu-id="7e02d-145">Title</span></span>

<span data-ttu-id="7e02d-146"> **xsi:type** が "ShowTaskpane" の場合に省略可能な要素。</span><span class="sxs-lookup"><span data-stu-id="7e02d-146">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="7e02d-147">この操作に関する、作業ウィンドウのカスタム タイトルを指定します。</span><span class="sxs-lookup"><span data-stu-id="7e02d-147">Specifies the custom title for the task pane for this action.</span></span>

<span data-ttu-id="7e02d-148">次の例は、 **Title**要素を使用するアクションを示しています。</span><span class="sxs-lookup"><span data-stu-id="7e02d-148">The following example shows an action that uses the **Title** element.</span></span> <span data-ttu-id="7e02d-149">**タイトル**を文字列に直接割り当てることはないことに注意してください。</span><span class="sxs-lookup"><span data-stu-id="7e02d-149">Note that you don't assign the **Title** to a string directly.</span></span> <span data-ttu-id="7e02d-150">代わりに、マニフェストの [**リソース**] セクションで定義されたリソース ID (resid) を割り当てます。</span><span class="sxs-lookup"><span data-stu-id="7e02d-150">Instead, you assign it a resource ID (resid), that is defined in the **Resources** section of the manifest.</span></span>

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

## <a name="supportspinning"></a><span data-ttu-id="7e02d-151">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="7e02d-151">SupportsPinning</span></span>

<span data-ttu-id="7e02d-152">**xsi:type** が "ShowTaskpane" の場合に省略可能な要素。</span><span class="sxs-lookup"><span data-stu-id="7e02d-152">Optional element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="7e02d-153">これを収容している [VersionOverrides](versionoverrides.md) 要素は、`xsi:type` 属性の値が `VersionOverridesV1_1` になっている必要があります。</span><span class="sxs-lookup"><span data-stu-id="7e02d-153">The containing [VersionOverrides](versionoverrides.md) elements must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span> <span data-ttu-id="7e02d-154">作業ウィンドウのピン留めをサポートする場合は、この要素に `true` の値を含めます。</span><span class="sxs-lookup"><span data-stu-id="7e02d-154">Include this element with a value of `true` to support task pane pinning.</span></span> <span data-ttu-id="7e02d-155">ユーザーは、作業ウィンドウをピン留めできるようになります。ピン留めすると、選択を変更したときも作業ウィンドウが開いたままになります。</span><span class="sxs-lookup"><span data-stu-id="7e02d-155">The user will be able to "pin" the task pane, causing it to stay open when changing the selection.</span></span> <span data-ttu-id="7e02d-156">詳細については、「[Outlook にピン留め可能な作業ウィンドウを実装する](../../outlook/pinnable-taskpane.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7e02d-156">For more information, see [Implement a pinnable task pane in Outlook](../../outlook/pinnable-taskpane.md).</span></span>

> [!NOTE]
> <span data-ttu-id="7e02d-157">サポートされている回転は、現在、Windows 2016 以降 (ビルド7628.1000 以降) および Outlook 2016 以降の Mac (ビルド16.13.503 以降) でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="7e02d-157">SupportsPinning is currently only supported by Outlook 2016 or later on Windows (build 7628.1000 or later) and Outlook 2016 or later on Mac (build 16.13.503 or later).</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```
