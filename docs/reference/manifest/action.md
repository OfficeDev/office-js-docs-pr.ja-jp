---
title: マニフェスト ファイルの Action 要素
description: この要素は、ユーザーがボタンまたはメニュー コントロールを選択するときに実行するアクションを指定します。
ms.date: 02/12/2021
localization_priority: Normal
ms.openlocfilehash: 6be1430800dea27dbd9bf78607161d88e475c145
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505410"
---
# <a name="action-element"></a><span data-ttu-id="8c7f4-103">Action 要素</span><span class="sxs-lookup"><span data-stu-id="8c7f4-103">Action element</span></span>

<span data-ttu-id="8c7f4-104">ユーザーが Button コントロールまたは Menu コントロールを選択するときに実行[するアクションを](control.md#button-control)[指定](control.md#menu-dropdown-button-controls)します。</span><span class="sxs-lookup"><span data-stu-id="8c7f4-104">Specifies the action to perform when the user selects a  [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) control.</span></span>

## <a name="attributes"></a><span data-ttu-id="8c7f4-105">属性</span><span class="sxs-lookup"><span data-stu-id="8c7f4-105">Attributes</span></span>

|  <span data-ttu-id="8c7f4-106">属性</span><span class="sxs-lookup"><span data-stu-id="8c7f4-106">Attribute</span></span>  |  <span data-ttu-id="8c7f4-107">必須</span><span class="sxs-lookup"><span data-stu-id="8c7f4-107">Required</span></span>  |  <span data-ttu-id="8c7f4-108">説明</span><span class="sxs-lookup"><span data-stu-id="8c7f4-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="8c7f4-109">xsi:type</span><span class="sxs-lookup"><span data-stu-id="8c7f4-109">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="8c7f4-110">はい</span><span class="sxs-lookup"><span data-stu-id="8c7f4-110">Yes</span></span>  | <span data-ttu-id="8c7f4-111">実行する操作の種類</span><span class="sxs-lookup"><span data-stu-id="8c7f4-111">Action type to take</span></span>|

## <a name="child-elements"></a><span data-ttu-id="8c7f4-112">子要素</span><span class="sxs-lookup"><span data-stu-id="8c7f4-112">Child elements</span></span>

|  <span data-ttu-id="8c7f4-113">要素</span><span class="sxs-lookup"><span data-stu-id="8c7f4-113">Element</span></span> |  <span data-ttu-id="8c7f4-114">説明</span><span class="sxs-lookup"><span data-stu-id="8c7f4-114">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="8c7f4-115">FunctionName</span><span class="sxs-lookup"><span data-stu-id="8c7f4-115">FunctionName</span></span>](#functionname) |    <span data-ttu-id="8c7f4-116">実行する関数の名前を指定します。</span><span class="sxs-lookup"><span data-stu-id="8c7f4-116">Specifies the name of the function to execute.</span></span> |
|  [<span data-ttu-id="8c7f4-117">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="8c7f4-117">SourceLocation</span></span>](#sourcelocation) |    <span data-ttu-id="8c7f4-118">この操作のソース ファイルの場所を指定します。</span><span class="sxs-lookup"><span data-stu-id="8c7f4-118">Specifies the source file location for this action.</span></span> |
|  [<span data-ttu-id="8c7f4-119">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="8c7f4-119">TaskpaneId</span></span>](#taskpaneid) | <span data-ttu-id="8c7f4-120">作業ウィンドウ コンテナーの ID を指定します。</span><span class="sxs-lookup"><span data-stu-id="8c7f4-120">Specifies the ID of the task pane container.</span></span>|
|  [<span data-ttu-id="8c7f4-121">Title</span><span class="sxs-lookup"><span data-stu-id="8c7f4-121">Title</span></span>](#title) | <span data-ttu-id="8c7f4-122">作業ウィンドウのカスタム タイトルを指定します。</span><span class="sxs-lookup"><span data-stu-id="8c7f4-122">Specifies the custom title for the task pane.</span></span>|
|  [<span data-ttu-id="8c7f4-123">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="8c7f4-123">SupportsPinning</span></span>](#supportspinning) | <span data-ttu-id="8c7f4-124">作業ウィンドウがピン留めをサポートすることを指定します。これにより、ユーザーが選択を変更したときも作業ウィンドウが開いたままになります。</span><span class="sxs-lookup"><span data-stu-id="8c7f4-124">Specifies that a task pane supports pinning, which keeps the task pane open when the user changes the selection.</span></span>|

## <a name="xsitype"></a><span data-ttu-id="8c7f4-125">xsi:type</span><span class="sxs-lookup"><span data-stu-id="8c7f4-125">xsi:type</span></span>

<span data-ttu-id="8c7f4-p101">この属性は、ユーザーがボタンをクリックしたときに実行される操作の種類を指定します。次のいずれかを指定できます。</span><span class="sxs-lookup"><span data-stu-id="8c7f4-p101">This attribute specifies the kind of action performed when the user selects the button. It can be one of the following:</span></span>

- `ExecuteFunction`
- `ShowTaskpane`

> [!IMPORTANT]
> <span data-ttu-id="8c7f4-128">**xsi:type**[が](../objectmodel/preview-requirement-set/office.context.mailbox.md#events)[.](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) `ExecuteFunction`</span><span class="sxs-lookup"><span data-stu-id="8c7f4-128">Registering [Mailbox](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) and [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) events is not available when **xsi:type** is `ExecuteFunction`.</span></span>

## <a name="functionname"></a><span data-ttu-id="8c7f4-129">FunctionName</span><span class="sxs-lookup"><span data-stu-id="8c7f4-129">FunctionName</span></span>

<span data-ttu-id="8c7f4-p102">**xsi:type** が "ExecuteFunction" のときに必ず指定する要素です。実行する関数の名前を指定します。関数は、[FunctionFile](functionfile.md) 要素に指定されたファイルに含まれています。</span><span class="sxs-lookup"><span data-stu-id="8c7f4-p102">Required element when **xsi:type** is "ExecuteFunction". Specifies the name of the function to execute. The function is contained in the file specified in the [FunctionFile](functionfile.md) element.</span></span>

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a><span data-ttu-id="8c7f4-133">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="8c7f4-133">SourceLocation</span></span>

<span data-ttu-id="8c7f4-134">**xsi:type が**"ShowTaskpane" の場合は必須の要素です。</span><span class="sxs-lookup"><span data-stu-id="8c7f4-134">Required element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="8c7f4-135">この操作のソース ファイルの場所を指定します。</span><span class="sxs-lookup"><span data-stu-id="8c7f4-135">Specifies the source file location for this action.</span></span> <span data-ttu-id="8c7f4-136">**resid 属性** は 32 文字以内で、Resources 要素の **Urls** 要素の **Url** 要素の **id** 属性の値に設定 [する必要](resources.md)があります。</span><span class="sxs-lookup"><span data-stu-id="8c7f4-136">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the [Resources](resources.md) element.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a><span data-ttu-id="8c7f4-137">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="8c7f4-137">TaskpaneId</span></span>

<span data-ttu-id="8c7f4-p104">**xsi:type** が "ShowTaskpane" の場合に省略可能な要素。作業ウィンドウ コンテナーの ID を指定します。複数の "ShowTaskpane" の操作があり、それぞれに対して独立したウィンドウを開く場合は、異なる **TaskpaneId** を使用します。同じウィンドウを共有する異なる操作に対しては、同じ **TaskpaneId** を使用します。ユーザーが同じ **TaskpaneId** を共有するコマンドを選択した場合、ウィンドウ コンテナーは開いたままですが、ウィンドウのコンテンツは対応する操作の "SourceLocation" に置き換えられます。</span><span class="sxs-lookup"><span data-stu-id="8c7f4-p104">Optional element when  **xsi:type** is "ShowTaskpane". Specifies the ID of the task pane container. When you have multiple "ShowTaskpane" actions, use a different **TaskpaneId** if you want an independent pane for each. Use the same **TaskpaneId** for  different actions that share the same pane. When users choose commands that share the same **TaskpaneId**, the pane container will remain open but the contents of the pane will be replaced with the corresponding Action "SourceLocation".</span></span>

> [!NOTE]
> <span data-ttu-id="8c7f4-143">この要素は、Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8c7f4-143">This element is not supported in Outlook.</span></span>

<span data-ttu-id="8c7f4-144">次の例では、同じ **TaskpaneId** を共有する 2 つのアクションを示します。</span><span class="sxs-lookup"><span data-stu-id="8c7f4-144">The following example shows two actions that share the same **TaskpaneId**.</span></span>

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

<span data-ttu-id="8c7f4-p105">次の例では、異なる **TaskpaneId** を使用する 2 つの操作を示します。これらの例を全体的な流れで確認する場合は、「[Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8c7f4-p105">The following examples show two actions that use a different **TaskpaneId**. To see these examples in context, see [Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span></span>

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

## <a name="title"></a><span data-ttu-id="8c7f4-147">役職</span><span class="sxs-lookup"><span data-stu-id="8c7f4-147">Title</span></span>

<span data-ttu-id="8c7f4-148">**xsi:type** が "ShowTaskpane" の場合に省略可能な要素。</span><span class="sxs-lookup"><span data-stu-id="8c7f4-148">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="8c7f4-149">この操作に関する、作業ウィンドウのカスタム タイトルを指定します。</span><span class="sxs-lookup"><span data-stu-id="8c7f4-149">Specifies the custom title for the task pane for this action.</span></span>

<span data-ttu-id="8c7f4-150">次の例は、Title 要素を使用するアクション **を示** しています。</span><span class="sxs-lookup"><span data-stu-id="8c7f4-150">The following example shows an action that uses the **Title** element.</span></span> <span data-ttu-id="8c7f4-151">Title を文字列に **直接割り** 当てない点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="8c7f4-151">Note that you don't assign the **Title** to a string directly.</span></span> <span data-ttu-id="8c7f4-152">代わりに、マニフェストの [リソース] セクションで定義されているリソース ID  (常駐) を割り当て、32 文字以内にできます。</span><span class="sxs-lookup"><span data-stu-id="8c7f4-152">Instead, you assign it a resource ID (resid), that is defined in the **Resources** section of the manifest and can be no more than 32 characters.</span></span>

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

## <a name="supportspinning"></a><span data-ttu-id="8c7f4-153">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="8c7f4-153">SupportsPinning</span></span>

<span data-ttu-id="8c7f4-154">**xsi:type** が "ShowTaskpane" の場合に省略可能な要素。</span><span class="sxs-lookup"><span data-stu-id="8c7f4-154">Optional element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="8c7f4-155">これを収容している [VersionOverrides](versionoverrides.md) 要素は、`xsi:type` 属性の値が `VersionOverridesV1_1` になっている必要があります。</span><span class="sxs-lookup"><span data-stu-id="8c7f4-155">The containing [VersionOverrides](versionoverrides.md) elements must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span> <span data-ttu-id="8c7f4-156">作業ウィンドウのピン留めをサポートする場合は、この要素に `true` の値を含めます。</span><span class="sxs-lookup"><span data-stu-id="8c7f4-156">Include this element with a value of `true` to support task pane pinning.</span></span> <span data-ttu-id="8c7f4-157">ユーザーは、作業ウィンドウをピン留めできるようになります。ピン留めすると、選択を変更したときも作業ウィンドウが開いたままになります。</span><span class="sxs-lookup"><span data-stu-id="8c7f4-157">The user will be able to "pin" the task pane, causing it to stay open when changing the selection.</span></span> <span data-ttu-id="8c7f4-158">詳細については、「[Outlook にピン留め可能な作業ウィンドウを実装する](../../outlook/pinnable-taskpane.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8c7f4-158">For more information, see [Implement a pinnable task pane in Outlook](../../outlook/pinnable-taskpane.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="8c7f4-159">要素は `SupportsPinning` 要件セット [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)で導入されましたが、現在サポートされているのは、以下を使用する Microsoft 365 サブスクライバーのみです。</span><span class="sxs-lookup"><span data-stu-id="8c7f4-159">Although the `SupportsPinning` element was introduced in [requirement set 1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), it's currently only supported for Microsoft 365 subscribers using the following.</span></span>
>
> - <span data-ttu-id="8c7f4-160">Windows 上の Outlook 2016 以降 (ビルド 7628.1000 以降)</span><span class="sxs-lookup"><span data-stu-id="8c7f4-160">Outlook 2016 or later on Windows (build 7628.1000 or later)</span></span>
> - <span data-ttu-id="8c7f4-161">Mac 上の Outlook 2016 以降 (ビルド 16.13.503 以降)</span><span class="sxs-lookup"><span data-stu-id="8c7f4-161">Outlook 2016 or later on Mac (build 16.13.503 or later)</span></span>
> - <span data-ttu-id="8c7f4-162">モダン Outlook on the web</span><span class="sxs-lookup"><span data-stu-id="8c7f4-162">Modern Outlook on the web</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```
