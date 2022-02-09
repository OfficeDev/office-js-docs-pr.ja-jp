---
title: マニフェスト ファイルの Button 型のコントロール要素
description: アクションを実行するか、作業ウィンドウを起動するボタンを定義します。
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: adc58424fe9898bffcbd9e16bed8f3b13b9df4a2
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467912"
---
# <a name="control-element-of-type-button"></a>Button 型のコントロール要素

アクションを実行するか、作業ウィンドウを起動するボタンを定義します。

> [!NOTE]
> この記事では、要素の属性に関する重要な情報を含む基本的な [Control](control.md) リファレンス記事に精通している必要があります。

ボタンは、ユーザーが選択したときに 1 つのアクションを実行します。 関数を実行するか、作業ウィンドウを表示します。 各ボタン コントロールには、マニフェスト内 `id` のすべての Control 要素間で一意の **属性値** が必要です。

> [!IMPORTANT]
> モバイル プラットフォームでは、"Button" 型コントロールは無視されます。 モバイル プラットフォームをサポートするには、"Button" 型のすべてのコントロールに対して"MobileButton" 型のコントロールを持っている必要があります。

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  [Label](#label)     | はい |  ボタンのテキストです。 |
|  **ToolTip**    |いいえ|ボタンのヒントです。 **resid 属性** は 32 文字以内で、String 要素の **id** 属性の値に設定する **必要** があります。 **String** 要素は、**LongStrings** 要素 ([Resources](resources.md) 要素の子要素) の子要素です。|
|  [Supertip](supertip.md)  | はい |  このボタンのヒントです。    |
|  [Icon](icon.md)      | はい |  ボタンの画像です。         |
|  [Action](action.md)    | はい |  実行するアクションを指定します。 Control 要素の Action 子 **は** 1 つ **のみです** 。 |
|  [Enabled (有効)](enabled.md)    | いいえ |  アドインの起動時にコントロールを有効にするかどうかを指定します。  |
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | いいえ |  カスタム コンテキスト タブをサポートするアプリケーションとプラットフォームの組み合わせにボタンを表示するかどうかを指定します。 使用する場合は、最初の子 *要素である* 必要があります。 |

### <a name="label"></a>ラベル

ボタンのテキストを、32 文字以内にできる唯一の属性を使用して指定し、Resources 要素の **ShortStrings** 子の **String** 要素の **id** 属性の値に設定する [必要](resources.md)があります。

**アドインの種類:** 作業ウィンドウ, メール

**次の VersionOverrides スキーマでのみ有効です**。

- 作業ウィンドウ 1.0
- メール 1.0
- メール 1.1

詳細については、「Version [overrides in the manifest」を参照してください](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) 親 **VersionOverrides が** Taskpane 1.0 と入力されている場合。
- 親 **VersionOverrides が Mail** 1.0 と入力されている場合のメールボックス [1.3](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)。
- 親 **VersionOverrides が Mail** 1.1 と入力されている場合のメールボックス [1.5](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)。

## <a name="examples"></a>例

次の例では、ボタンは関数を実行します。 また、アドインの起動時に無効にするように構成されています。 プログラムで有効にできます。 詳細については、「[アドイン コマンドを有効または無効にする](../../design/disable-add-in-commands.md)」を参照してください。

```xml
<Control xsi:type="Button" id="Contoso.msgReadFunctionButton">
  <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
  <Label resid="funcReadButtonLabel" />
  <Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="blue-icon-16" />
    <bt:Image size="32" resid="blue-icon-32" />
    <bt:Image size="80" resid="blue-icon-80" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
  </Action>
  <Enabled>false</Enabled>
</Control>
```

次の例では、ボタンに作業ウィンドウが表示されます。

```xml
<Control xsi:type="Button" id="Contoso.msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Supertip>
    <Title resid="paneReadSuperTipTitle" />
    <Description resid="paneReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="green-icon-16" />
    <bt:Image size="32" resid="green-icon-32" />
    <bt:Image size="80" resid="green-icon-80" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
</Control>
```
