---
title: Outlook アドインにモバイル サポートを追加する
description: Outlook Mobile のサポートを追加するには、アドイン マニフェストを更新する必要があります。さらに、モバイル シナリオのコードを変更することが必要な場合もあります。
ms.date: 07/16/2021
localization_priority: Normal
ms.openlocfilehash: afd8c944a6524f4b5b532ed6136ef28c326d0d0c475b35d6254592c3a392aac3
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57089541"
---
# <a name="add-support-for-add-in-commands-for-outlook-mobile"></a>Outlook Mobile のアドイン コマンドのサポートを追加する

Outlook Mobile でアドイン コマンドを使用すると、ユーザーは Outlook on the web、Windows、および Mac で既に持っているのと同じ機能 (いくつかの制限[があります)](#code-considerations)にアクセスできます。 Outlook Mobile のサポートを追加するには、アドイン マニフェストを更新する必要があります。さらに、モバイル シナリオのコードを変更することが必要な場合もあります。

## <a name="updating-the-manifest"></a>マニフェストを更新する

Outlook Mobile でアドイン コマンドを有効にするための最初の手順は、アドイン マニフェストでの定義です。[VersionOverrides](../reference/manifest/versionoverrides.md) v1.1 スキーマは、モバイル用に新しいフォーム ファクター [MobileFormFactor](../reference/manifest/mobileformfactor.md) を定義します。

この要素には、モバイル クライアントにアドインを読み込むためのすべての情報が含まれています。これにより、モバイル エクスペリエンスに対して完全に異なる UI 要素と JavaScript ファイルを定義することができます。

次の使用例は、要素内の 1 つの作業ウィンドウ ボタンを示 `MobileFormFactor` しています。

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

メールボックス要件セット 1.5 では、REST API と互換性のあるアクセス トークンを要求できる新しいバージョンの[Office.context.mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)と、ユーザーの REST API エンドポイントの検索に使用できる新しい[Office.context.mailbox.restUrl](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties)プロパティが導入されました。

### <a name="pinch-zoom"></a>ピンチによるズーム

既定で、ユーザーは "ピンチによるズーム" ジェスチャを使用して作業ウィンドウで拡大することができます。ご使用のシナリオでこれが該当しない場合は、HTML でピンチによるズームを無効にしてください。

### <a name="close-task-panes"></a>作業ウィンドウを閉じる

Outlook Mobile では、作業ウィンドウが画面全体を占めるので、既定ではユーザーが作業ウィンドウを閉じてメッセージに戻る必要があります。シナリオが完成したら、[Office.context.ui.closeContainer](/javascript/api/office/office.ui#closeContainer__) メソッドを使用して作業ウィンドウを閉じることを検討してください。

### <a name="compose-mode-and-appointments"></a>作成モードと予定

現在、Outlook Mobile のアドインは、メッセージの読み取り時にのみアクティブ化をサポートしています。 メッセージを作成するときや、予定を表示または作成するときには、アドインはアクティブ化されません。 ただし、オンライン会議プロバイダー統合アドインは、予定オーガナイザー モードでアクティブ化できます。 この例外 (使用可能な API を含む) の詳細については、「オンライン会議プロバイダー Outlookを作成する」[を参照してください](online-meeting.md#available-apis)。

### <a name="unsupported-apis"></a>サポートされていない API

要件セット 1.6 以降で導入された API は、モバイル ではOutlookされません。 以前の要件セットの次の API もサポートされていません。

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

[Exchange サーバーと Outlook クライアントでサポートされる要件セット](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)