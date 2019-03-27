---
title: マニフェスト ファイルの Action 要素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 59df6cce6af1277f365a1dd3cd0b3ef11230804e
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870829"
---
# <a name="action-element"></a>Action 要素

ユーザーが[ボタン](control.md#button-control)または[メニュー](control.md#menu-dropdown-button-controls) コントロールを選択したときに実行する操作を指定します。

## <a name="attributes"></a>属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  はい  | 実行する操作の種類|

## <a name="child-elements"></a>子要素

|  要素 |  説明  |
|:-----|:-----|
|  [FunctionName](#functionname) |    実行する関数の名前を指定します。 |
|  [SourceLocation](#sourcelocation) |    この操作のソース ファイルの場所を指定します。 |
|  [TaskpaneId](#taskpaneid) | 作業ウィンドウ コンテナーの ID を指定します。|
|  [Title](#title) | 作業ウィンドウのカスタム タイトルを指定します。|
|  [SupportsPinning](#supportspinning) | 作業ウィンドウがピン留めをサポートすることを指定します。これにより、ユーザーが選択を変更したときも作業ウィンドウが開いたままになります。|
  

## <a name="xsitype"></a>xsi:type

この属性は、ユーザーがボタンをクリックしたときに実行される操作の種類を指定します。次のいずれかを指定できます。

- `ExecuteFunction`
- `ShowTaskpane`

## <a name="functionname"></a>FunctionName

**xsi:type** が "ExecuteFunction" のときに必ず指定する要素です。実行する関数の名前を指定します。関数は、[FunctionFile](functionfile.md) 要素に指定されたファイルに含まれています。

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a>SourceLocation

**xsi:type** が "ShowTaskpane" のときに必ず指定する要素です。このアクションのソース ファイルの場所を指定します。 **resid** 属性は、 **Resources** 要素の **Urls** 要素にある **Url** 要素の [id](resources.md) 属性の値を指定します。

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a>TaskpaneId

 **xsi:type** が "ShowTaskpane" の場合に省略可能な要素。 作業ウィンドウ コンテナーの ID を指定します。 複数の "ShowTaskpane" の操作があり、それぞれに対して独立したウィンドウを開く場合は、異なる **TaskpaneId** を使用します。 同じウィンドウを共有する異なる操作に対しては、同じ **TaskpaneId** を使用します。 ユーザーが同じ **TaskpaneId** を共有するコマンドを選択した場合、ウィンドウ コンテナーは開いたままですが、ウィンドウのコンテンツは対応する Action "SourceLocation" に置き換えられます。

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

 **xsi:type** が "ShowTaskpane" の場合に省略可能な要素。 この操作に関する、作業ウィンドウのカスタム タイトルを指定します。

以下の例は、**Title** 要素を使用する 2 つの異なるアクションを示します。

```xml
<Action xsi:type="ShowTaskpane">
<TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
<SourceLocation resid="PG.Code.Url" />
<Title resid="PG.CodeCommand.Title" />
</Action>
```

```xml
<Action xsi:type="ShowTaskpane">
<SourceLocation resid="PG.Run.Url" />
<Title resid="PG.RunCommand.Title" />
</Action>
```

```xml
<bt:Urls>
<bt:Url id="PG.Code.Url" DefaultValue="https://localhost:3000?commands=1" />
<bt:Url id="PG.Run.Url" DefaultValue="https://localhost:3000/run.html" />
</bt:Urls>
```

```xml
<bt:ShortStrings>
<bt:String id="PG.CodeCommand.Title" DefaultValue="Code" />
<bt:String id="PG.RunCommand.Title" DefaultValue="Run" />
</bt:ShortStrings>
```

## <a name="supportspinning"></a>SupportsPinning

**xsi:type** が "ShowTaskpane" の場合に省略可能な要素。 これを収容している [VersionOverrides](versionoverrides.md) 要素は、`xsi:type` 属性の値が `VersionOverridesV1_1` になっている必要があります。 作業ウィンドウのピン留めをサポートする場合は、この要素に `true` の値を含めます。 ユーザーは、作業ウィンドウをピン留めできるようになります。ピン留めすると、選択を変更したときも作業ウィンドウが開いたままになります。 詳細については、「[Outlook にピン留め可能な作業ウィンドウを実装する](/outlook/add-ins/pinnable-taskpane)」を参照してください。

> [!NOTE]
> サポートされている回転は、現在、outlook 2016 for Windows (ビルド7628.1000 以降) と outlook 2016 for Mac (ビルド16.13.503 以降) でのみサポートされています。

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```
