---
title: マニフェスト ファイルの Action 要素
description: この要素は、ユーザーがボタンまたはメニュー コントロールを選択するときに実行するアクションを指定します。
ms.date: 02/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 21c8f9a6345641f23aad70efed67c9c45f72a1c8
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340415"
---
# <a name="action-element"></a>Action 要素

ユーザーが Button コントロールまたは Menu コントロールを選択するときに実行[するアクションを](control-button.md)[指定](control-menu.md)します。

**次の VersionOverrides スキーマでのみ有効です**。

- 作業ウィンドウ 1.0
- メール 1.0
- メール 1.1

詳細については、「Version [overrides in the manifest」を参照してください](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) 親 **VersionOverrides が** Taskpane 1.0 と入力されている場合。
- 親 **VersionOverrides が Mail** 1.0 と入力されている場合のメールボックス [1.3](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)。
- 親 **VersionOverrides が Mail** 1.1 と入力されている場合のメールボックス [1.5](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)。

## <a name="attributes"></a>属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  はい  | 実行する操作の種類|

## <a name="child-elements"></a>子要素

|  要素 |  説明  |
|:-----|:-----|
|  [FunctionName](#functionname) |    実行する関数の名前を指定します。 |
|  [SourceLocation](#sourcelocation) |    この操作のソース ファイルの場所を指定します。 |
|  [TaskpaneId](#taskpaneid) | 作業ウィンドウ コンテナーの ID を指定します。 このアドインではOutlookサポートされていません。|
|  [Title](#title) | 作業ウィンドウのカスタム タイトルを指定します。 このアドインではOutlookサポートされていません。|
|  [SupportsPinning](#supportspinning) | 作業ウィンドウがピン留めをサポートすることを指定します。これにより、ユーザーが選択を変更したときも作業ウィンドウが開いたままになります。|

## <a name="xsitype"></a>xsi:type

この属性は、ユーザーがボタンをクリックしたときに実行される操作の種類を指定します。次のいずれかを指定できます。

- `ExecuteFunction`
- `ShowTaskpane`

> [!IMPORTANT]
> **xsi:type** `ExecuteFunction`[が](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) [.](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events)

## <a name="functionname"></a>FunctionName

**xsi:type** `ExecuteFunction`が . 実行する関数の名前を指定します。 関数は、[FunctionFile](functionfile.md) 要素に指定されたファイルに含まれています。

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a>SourceLocation

**xsi:type** `ShowTaskpane`が . この操作のソース ファイルの場所を指定します。 **resid 属性** は 32 文字以内で、Resources 要素の **Urls** 要素の **Url** 要素の **id** 属性の値に設定 [する必要](resources.md)があります。

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a>TaskpaneId

**xsi:type** `ShowTaskpane`が . 作業ウィンドウ コンテナーの ID を指定します。 複数のアクションがある場合 `ShowTaskpane` は、それぞれ独立したウィンドウが必要な場合は、別の **TaskpaneId** を使用します。 同じウィンドウを共有する異なる操作に対しては、同じ **TaskpaneId** を使用します。 ユーザーが同じ **TaskpaneId** を共有するコマンドを選択すると、ウィンドウ コンテナーは開いたままですが、ウィンドウの内容は対応する Action に置き換えます `SourceLocation`。

**アドインの種類:** 作業ウィンドウ

**次の VersionOverrides スキーマでのみ有効です**。

- 作業ウィンドウ 1.0

詳細については、「Version [overrides in the manifest」を参照してください](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md)

> [!NOTE]
> この要素は、Outlook ではサポートされていません。

次の例では、同じ **TaskpaneId** を共有する 2 つのアクションを示します。

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

次の例では、異なる **TaskpaneId** を使用する 2 つの操作を示します。これらの例を全体的な流れで確認する場合は、「[Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml)」を参照してください。

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

## <a name="title"></a>役職

**xsi:type** `ShowTaskpane`が . この操作に関する、作業ウィンドウのカスタム タイトルを指定します。

**アドインの種類:** 作業ウィンドウ

**次の VersionOverrides スキーマでのみ有効です**。

- 作業ウィンドウ 1.0

詳細については、「Version [overrides in the manifest」を参照してください](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md)

> [!NOTE]
> この子要素は、アドインOutlookサポートされていません。

次の例は、Title 要素を使用するアクション **を示** しています。 Title を文字列に **直接割り** 当てない点に注意してください。 代わりに、マニフェストの [リソース] セクションで定義されているリソース ID (常駐)  を割り当て、32 文字以内にできます。

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

## <a name="supportspinning"></a>SupportsPinning

**xsi:type** `ShowTaskpane`が . 含まれている [VersionOverrides 要素](versionoverrides.md) には **、xsi:type 属性値が** 必要です `VersionOverridesV1_1`。 作業ウィンドウのピン留めをサポートする場合は、この要素に `true` の値を含めます。 ユーザーは、作業ウィンドウをピン留めできるようになります。ピン留めすると、選択を変更したときも作業ウィンドウが開いたままになります。 詳細については、「[Outlook にピン留め可能な作業ウィンドウを実装する](../../outlook/pinnable-taskpane.md)」を参照してください。

**アドインの種類:** メール

**次の VersionOverrides スキーマでのみ有効です**。

- メール 1.1

詳細については、「Version [overrides in the manifest」を参照してください](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [Mailbox 1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)

> [!IMPORTANT]
> **SupportPinning 要素** は要件セット [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) で導入されましたが、現在サポートされているのは、以下を使用する Microsoft 365サブスクライバーのみです。
>
> - Outlook 2016以降 (ビルド 7628.1000 以降) Windows(ビルド 7628.1000 以降)
> - Outlook 2016以降の Mac (ビルド 16.13.503 以降)
> - モダン Outlook on the web

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```
