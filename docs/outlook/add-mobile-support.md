---
title: Outlook アドインにモバイル サポートを追加する
description: Outlook Mobile のサポートを追加するには、アドイン マニフェストを更新する必要があります。さらに、モバイル シナリオのコードを変更することが必要な場合もあります。
ms.date: 12/10/2019
localization_priority: Normal
ms.openlocfilehash: 31f58102129ae207da55839f7b48cc8a060645ad
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720856"
---
# <a name="add-support-for-add-in-commands-for-outlook-mobile"></a>Outlook Mobile のアドイン コマンドのサポートを追加する

Outlook Mobile でアドインコマンドを使用すると、ユーザーは web 上の Outlook、Windows、および Mac で既に所有しているものと同じ機能 (一部の[制限](#code-considerations)あり) にアクセスできます。 Outlook Mobile のサポートを追加するには、アドイン マニフェストを更新する必要があります。さらに、モバイル シナリオのコードを変更することが必要な場合もあります。

## <a name="updating-the-manifest"></a>マニフェストを更新する

Outlook Mobile でアドイン コマンドを有効にするための最初の手順は、アドイン マニフェストでの定義です。[VersionOverrides](../reference/manifest/versionoverrides.md) v1.1 スキーマは、モバイル用に新しいフォーム ファクター [MobileFormFactor](../reference/manifest/mobileformfactor.md) を定義します。

この要素には、モバイル クライアントにアドインを読み込むためのすべての情報が含まれています。これにより、モバイル エクスペリエンスに対して完全に異なる UI 要素と JavaScript ファイルを定義することができます。

次の例は、 `MobileFormFactor`要素内の1つの作業ウィンドウボタンを示しています。

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
  ...
  <MobileFormFactor>
    <FunctionFile resid="residUILessFunctionFileUrl" />
    <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
      <Group id="mobileMsgRead">
        <Label resid="groupLabel" />
        <Control xsi:type="MobileButton" id="TaskPaneBtn">
          <Label resid="residTaskPaneButtonName" />
          <Icon xsi:type="bt:MobileIconList">
            <bt:Image size="25" scale="1" resid="tp0icon" />
            <bt:Image size="25" scale="2" resid="tp0icon" />
            <bt:Image size="25" scale="3" resid="tp0icon" />

            <bt:Image size="32" scale="1" resid="tp0icon" />
            <bt:Image size="32" scale="2" resid="tp0icon" />
            <bt:Image size="32" scale="3" resid="tp0icon" />

            <bt:Image size="48" scale="1" resid="tp0icon" />
            <bt:Image size="48" scale="2" resid="tp0icon" />
            <bt:Image size="48" scale="3" resid="tp0icon" />
          </Icon>
          <Action xsi:type="ShowTaskpane">
            <SourceLocation resid="residTaskpaneUrl" />
          </Action>
        </Control>
      </Group>
    </ExtensionPoint>
  </MobileFormFactor>
  ...
</VersionOverrides>
```

これは、[DesktopFormFactor](../reference/manifest/desktopformfactor.md) 要素に表示される要素と非常によく似ていますが、いくつかの注目すべき違いがあります。

- [OfficeTab](../reference/manifest/officetab.md) 要素は使用されません。
- [ExtensionPoint](../reference/manifest/extensionpoint.md) 要素に含まれる子要素は 1 つでなければなりません。アドインがボタンを 1 つのみ追加する場合、子要素は [Control](../reference/manifest/control.md) 要素になります。アドインがボタンを複数追加する場合、子要素は複数の `Control` 要素を含む [Group](../reference/manifest/group.md) 要素になります。
- `Control` 要素に相当する `Menu` の種類はありません。
- [Supertip](../reference/manifest/supertip.md) 要素は使用されません。
- アイコンの必須サイズが異なります。モバイル アドインは少なくとも 25x25、32x32 および 48x48 ピクセルのアイコンをサポートする必要があります。

## <a name="code-considerations"></a>コードに関する考慮事項

モバイル用のアドインの設計には、追加の考慮事項がいくつか導入されています。

### <a name="use-rest-instead-of-exchange-web-services"></a>Exchange Web サービスの代わりに REST を使用する

[Office.context.mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) メソッドは、Outlook Mobile ではサポートされていません。可能な場合には、アドインは優先的に Office.js API から情報を取得します。Office.js API によって表示されていない情報がアドインで必要な場合、[Outlook REST APIs](/outlook/rest/) を使用してユーザーのメールボックスにアクセスする必要があります。

メールボックス要件セット1.5 には、REST Api と互換性のあるアクセストークンを要求できる新しいバージョンの[office.context.mailbox.resturl が](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties)プロパティと、ユーザーの rest api エンドポイントを検索するために使用できる新しいバージョンのプロパティが導入[されて](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)います。

### <a name="pinch-zoom"></a>ピンチによるズーム

既定で、ユーザーは "ピンチによるズーム" ジェスチャを使用して作業ウィンドウで拡大することができます。ご使用のシナリオでこれが該当しない場合は、HTML でピンチによるズームを無効にしてください。

### <a name="close-task-panes"></a>作業ウィンドウを閉じる

Outlook Mobile では、作業ウィンドウが画面全体を占めるので、既定ではユーザーが作業ウィンドウを閉じてメッセージに戻る必要があります。シナリオが完成したら、[Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--) メソッドを使用して作業ウィンドウを閉じることを検討してください。

### <a name="compose-mode-and-appointments"></a>作成モードと予定

現在、Outlook Mobile のアドインは、メッセージ読み取り時のアクティブ化のみをサポートしています。メッセージを作成するときや、予定を表示または作成するときには、アドインはアクティブ化されません。

### <a name="unsupported-apis"></a>サポートされていない API

要件セット1.6 以降で導入された Api は、Outlook Mobile ではサポートされていません。 以前の要件セットからの次の Api もサポートされていません。

  - [Office.context.officeTheme](../reference/objectmodel/preview-requirement-set/office.context.md#officetheme-officetheme)
  - [Office.context.mailbox.ewsUrl](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties)
  - [Office.context.mailbox.convertToEwsId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [Office.context.mailbox.convertToRestId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [Office.context.mailbox.displayAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [Office.context.mailbox.displayMessageForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [Office.context.mailbox.displayNewAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [Office.context.mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [Office.context.mailbox.item.dateTimeModified](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
  - [Office.context.mailbox.item.displayReplyAllForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [Office.context.mailbox.item.displayReplyForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [Office.context.mailbox.item.getEntities](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [Office.context.mailbox.item.getEntitiesByType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [Office.context.mailbox.item.getFilteredEntitiesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [Office.context.mailbox.item.getRegexMatches](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [Office.context.mailbox.item.getRegexMatchesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)

## <a name="see-also"></a>関連項目

[要件セットのサポート](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)