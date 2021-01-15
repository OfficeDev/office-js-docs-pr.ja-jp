---
title: マニフェスト ファイルの Action 要素
description: この要素は、ユーザーがボタンまたはメニュー コントロールを選択するときに実行するアクションを指定します。
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: e345d0a1682e0125373a309e1e56eb2d6298ac7d
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771415"
---
# <a name="action-element"></a><span data-ttu-id="e8b35-103">Action 要素</span><span class="sxs-lookup"><span data-stu-id="e8b35-103">Action element</span></span>

<span data-ttu-id="e8b35-104">ユーザーがボタン コントロールまたはメニュー コントロールを選択するときに実行  [するアクション](control.md#button-control) を [指定](control.md#menu-dropdown-button-controls) します。</span><span class="sxs-lookup"><span data-stu-id="e8b35-104">Specifies the action to perform when the user selects a  [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) control.</span></span>

## <a name="attributes"></a><span data-ttu-id="e8b35-105">属性</span><span class="sxs-lookup"><span data-stu-id="e8b35-105">Attributes</span></span>

|  <span data-ttu-id="e8b35-106">属性</span><span class="sxs-lookup"><span data-stu-id="e8b35-106">Attribute</span></span>  |  <span data-ttu-id="e8b35-107">必須</span><span class="sxs-lookup"><span data-stu-id="e8b35-107">Required</span></span>  |  <span data-ttu-id="e8b35-108">説明</span><span class="sxs-lookup"><span data-stu-id="e8b35-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="e8b35-109">xsi:type</span><span class="sxs-lookup"><span data-stu-id="e8b35-109">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="e8b35-110">はい</span><span class="sxs-lookup"><span data-stu-id="e8b35-110">Yes</span></span>  | <span data-ttu-id="e8b35-111">実行する操作の種類</span><span class="sxs-lookup"><span data-stu-id="e8b35-111">Action type to take</span></span>|

## <a name="child-elements"></a><span data-ttu-id="e8b35-112">子要素</span><span class="sxs-lookup"><span data-stu-id="e8b35-112">Child elements</span></span>

|  <span data-ttu-id="e8b35-113">要素</span><span class="sxs-lookup"><span data-stu-id="e8b35-113">Element</span></span> |  <span data-ttu-id="e8b35-114">説明</span><span class="sxs-lookup"><span data-stu-id="e8b35-114">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="e8b35-115">FunctionName</span><span class="sxs-lookup"><span data-stu-id="e8b35-115">FunctionName</span></span>](#functionname) |    <span data-ttu-id="e8b35-116">実行する関数の名前を指定します。</span><span class="sxs-lookup"><span data-stu-id="e8b35-116">Specifies the name of the function to execute.</span></span> |
|  [<span data-ttu-id="e8b35-117">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="e8b35-117">SourceLocation</span></span>](#sourcelocation) |    <span data-ttu-id="e8b35-118">この操作のソース ファイルの場所を指定します。</span><span class="sxs-lookup"><span data-stu-id="e8b35-118">Specifies the source file location for this action.</span></span> |
|  [<span data-ttu-id="e8b35-119">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="e8b35-119">TaskpaneId</span></span>](#taskpaneid) | <span data-ttu-id="e8b35-120">作業ウィンドウ コンテナーの ID を指定します。</span><span class="sxs-lookup"><span data-stu-id="e8b35-120">Specifies the ID of the task pane container.</span></span>|
|  [<span data-ttu-id="e8b35-121">Title</span><span class="sxs-lookup"><span data-stu-id="e8b35-121">Title</span></span>](#title) | <span data-ttu-id="e8b35-122">作業ウィンドウのカスタム タイトルを指定します。</span><span class="sxs-lookup"><span data-stu-id="e8b35-122">Specifies the custom title for the task pane.</span></span>|
|  [<span data-ttu-id="e8b35-123">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="e8b35-123">SupportsPinning</span></span>](#supportspinning) | <span data-ttu-id="e8b35-124">作業ウィンドウがピン留めをサポートすることを指定します。これにより、ユーザーが選択を変更したときも作業ウィンドウが開いたままになります。</span><span class="sxs-lookup"><span data-stu-id="e8b35-124">Specifies that a task pane supports pinning, which keeps the task pane open when the user changes the selection.</span></span>|
  

## <a name="xsitype"></a><span data-ttu-id="e8b35-125">xsi:type</span><span class="sxs-lookup"><span data-stu-id="e8b35-125">xsi:type</span></span>

<span data-ttu-id="e8b35-p101">この属性は、ユーザーがボタンをクリックしたときに実行される操作の種類を指定します。次のいずれかを指定できます。</span><span class="sxs-lookup"><span data-stu-id="e8b35-p101">This attribute specifies the kind of action performed when the user selects the button. It can be one of the following:</span></span>

- `ExecuteFunction`
- `ShowTaskpane`

## <a name="functionname"></a><span data-ttu-id="e8b35-128">FunctionName</span><span class="sxs-lookup"><span data-stu-id="e8b35-128">FunctionName</span></span>

<span data-ttu-id="e8b35-p102">**xsi:type** が "ExecuteFunction" のときに必ず指定する要素です。実行する関数の名前を指定します。関数は、[FunctionFile](functionfile.md) 要素に指定されたファイルに含まれています。</span><span class="sxs-lookup"><span data-stu-id="e8b35-p102">Required element when **xsi:type** is "ExecuteFunction". Specifies the name of the function to execute. The function is contained in the file specified in the [FunctionFile](functionfile.md) element.</span></span>

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a><span data-ttu-id="e8b35-132">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="e8b35-132">SourceLocation</span></span>

<span data-ttu-id="e8b35-133">**xsi:type が**"ShowTaskpane" の場合は必須要素です。</span><span class="sxs-lookup"><span data-stu-id="e8b35-133">Required element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="e8b35-134">この操作のソース ファイルの場所を指定します。</span><span class="sxs-lookup"><span data-stu-id="e8b35-134">Specifies the source file location for this action.</span></span> <span data-ttu-id="e8b35-135">**resid 属性** は 32 文字以内で [、Resources](resources.md)要素の **Urls** 要素の **Url** 要素の **id** 属性の値に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e8b35-135">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the [Resources](resources.md) element.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a><span data-ttu-id="e8b35-136">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="e8b35-136">TaskpaneId</span></span>

<span data-ttu-id="e8b35-p104">**xsi:type** が "ShowTaskpane" の場合に省略可能な要素。作業ウィンドウ コンテナーの ID を指定します。複数の "ShowTaskpane" の操作があり、それぞれに対して独立したウィンドウを開く場合は、異なる **TaskpaneId** を使用します。同じウィンドウを共有する異なる操作に対しては、同じ **TaskpaneId** を使用します。ユーザーが同じ **TaskpaneId** を共有するコマンドを選択した場合、ウィンドウ コンテナーは開いたままですが、ウィンドウのコンテンツは対応する操作の "SourceLocation" に置き換えられます。</span><span class="sxs-lookup"><span data-stu-id="e8b35-p104">Optional element when  **xsi:type** is "ShowTaskpane". Specifies the ID of the task pane container. When you have multiple "ShowTaskpane" actions, use a different **TaskpaneId** if you want an independent pane for each. Use the same **TaskpaneId** for  different actions that share the same pane. When users choose commands that share the same **TaskpaneId**, the pane container will remain open but the contents of the pane will be replaced with the corresponding Action "SourceLocation".</span></span>

> [!NOTE]
> <span data-ttu-id="e8b35-142">この要素は、Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e8b35-142">This element is not supported in Outlook.</span></span>

<span data-ttu-id="e8b35-143">次の例では、同じ **TaskpaneId** を共有する 2 つのアクションを示します。</span><span class="sxs-lookup"><span data-stu-id="e8b35-143">The following example shows two actions that share the same **TaskpaneId**.</span></span>

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

<span data-ttu-id="e8b35-p105">次の例では、異なる **TaskpaneId** を使用する 2 つの操作を示します。これらの例を全体的な流れで確認する場合は、「[Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e8b35-p105">The following examples show two actions that use a different **TaskpaneId**. To see these examples in context, see [Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span></span>

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

## <a name="title"></a><span data-ttu-id="e8b35-146">役職</span><span class="sxs-lookup"><span data-stu-id="e8b35-146">Title</span></span>

<span data-ttu-id="e8b35-147">**xsi:type** が "ShowTaskpane" の場合に省略可能な要素。</span><span class="sxs-lookup"><span data-stu-id="e8b35-147">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="e8b35-148">この操作に関する、作業ウィンドウのカスタム タイトルを指定します。</span><span class="sxs-lookup"><span data-stu-id="e8b35-148">Specifies the custom title for the task pane for this action.</span></span>

<span data-ttu-id="e8b35-149">次の例は、Title 要素を使用するアクション **を示** しています。</span><span class="sxs-lookup"><span data-stu-id="e8b35-149">The following example shows an action that uses the **Title** element.</span></span> <span data-ttu-id="e8b35-150">タイトルを文字列に **直接割り** 当てない点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="e8b35-150">Note that you don't assign the **Title** to a string directly.</span></span> <span data-ttu-id="e8b35-151">代わりに、マニフェストの [リソース] セクションで定義されているリソース ID  (resid) を割り当て、32 文字以下にできます。</span><span class="sxs-lookup"><span data-stu-id="e8b35-151">Instead, you assign it a resource ID (resid), that is defined in the **Resources** section of the manifest and can be no more than 32 characters.</span></span>

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

## <a name="supportspinning"></a><span data-ttu-id="e8b35-152">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="e8b35-152">SupportsPinning</span></span>

<span data-ttu-id="e8b35-153">**xsi:type** が "ShowTaskpane" の場合に省略可能な要素。</span><span class="sxs-lookup"><span data-stu-id="e8b35-153">Optional element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="e8b35-154">これを収容している [VersionOverrides](versionoverrides.md) 要素は、`xsi:type` 属性の値が `VersionOverridesV1_1` になっている必要があります。</span><span class="sxs-lookup"><span data-stu-id="e8b35-154">The containing [VersionOverrides](versionoverrides.md) elements must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span> <span data-ttu-id="e8b35-155">作業ウィンドウのピン留めをサポートする場合は、この要素に `true` の値を含めます。</span><span class="sxs-lookup"><span data-stu-id="e8b35-155">Include this element with a value of `true` to support task pane pinning.</span></span> <span data-ttu-id="e8b35-156">ユーザーは、作業ウィンドウをピン留めできるようになります。ピン留めすると、選択を変更したときも作業ウィンドウが開いたままになります。</span><span class="sxs-lookup"><span data-stu-id="e8b35-156">The user will be able to "pin" the task pane, causing it to stay open when changing the selection.</span></span> <span data-ttu-id="e8b35-157">詳細については、「[Outlook にピン留め可能な作業ウィンドウを実装する](../../outlook/pinnable-taskpane.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e8b35-157">For more information, see [Implement a pinnable task pane in Outlook](../../outlook/pinnable-taskpane.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="e8b35-158">この要素 `SupportsPinning` は要件セット [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)で導入されましたが、現在サポートされているのは、次を使用する Microsoft 365 サブスクライバーのみです。</span><span class="sxs-lookup"><span data-stu-id="e8b35-158">Although the `SupportsPinning` element was introduced in [requirement set 1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), it's currently only supported for Microsoft 365 subscribers using the following.</span></span>
> - <span data-ttu-id="e8b35-159">Windows 上の Outlook 2016 以降 (ビルド 7628.1000 以降)</span><span class="sxs-lookup"><span data-stu-id="e8b35-159">Outlook 2016 or later on Windows (build 7628.1000 or later)</span></span>
> - <span data-ttu-id="e8b35-160">Mac 上の Outlook 2016 以降 (ビルド 16.13.503 以降)</span><span class="sxs-lookup"><span data-stu-id="e8b35-160">Outlook 2016 or later on Mac (build 16.13.503 or later)</span></span>
> - <span data-ttu-id="e8b35-161">モダン Outlook on the web</span><span class="sxs-lookup"><span data-stu-id="e8b35-161">Modern Outlook on the web</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```
