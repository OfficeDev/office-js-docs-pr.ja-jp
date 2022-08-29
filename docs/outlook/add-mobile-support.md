---
title: Outlook アドインにモバイル サポートを追加する
description: 必要に応じて、アドイン マニフェストを更新し、モバイル シナリオのコードを変更する方法など、Outlook Mobile のサポートを追加する方法について説明します。
ms.date: 04/15/2022
ms.localizationpriority: medium
ms.openlocfilehash: 50f1613e83d9b23178714cfb3da8110a4c561b05
ms.sourcegitcommit: 57258dd38507f791bbb39cbb01d6bbd5a9d226b9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/12/2022
ms.locfileid: "67318880"
---
# <a name="add-support-for-add-in-commands-for-outlook-mobile"></a>Outlook Mobile のアドイン コマンドのサポートを追加する

Outlook Mobile でアドイン コマンドを使用すると、ユーザーは、Outlook on the web、Windows、Mac で既に持っているのと同じ機能 (いくつかの[制限](#code-considerations)あり) にアクセスできます。 Outlook Mobile のサポートを追加するには、アドイン マニフェストを更新する必要があります。さらに、モバイル シナリオのコードを変更することが必要な場合もあります。

## <a name="updating-the-manifest"></a>マニフェストを更新する

Outlook Mobile でアドイン コマンドを有効にするための最初の手順は、アドイン マニフェストでの定義です。[VersionOverrides](/javascript/api/manifest/versionoverrides) v1.1 スキーマは、モバイル用に新しいフォーム ファクター [MobileFormFactor](/javascript/api/manifest/mobileformfactor) を定義します。

この要素には、モバイル クライアントにアドインを読み込むためのすべての情報が含まれています。これにより、モバイル エクスペリエンスに対して完全に異なる UI 要素と JavaScript ファイルを定義することができます。

次の例は、要素内の 1 つの作業ウィンドウ ボタンを `MobileFormFactor` 示しています。

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

これは、[DesktopFormFactor](/javascript/api/manifest/desktopformfactor) 要素に表示される要素と非常によく似ていますが、いくつかの注目すべき違いがあります。

- [OfficeTab](/javascript/api/manifest/officetab) 要素は使用されません。
- [ExtensionPoint](/javascript/api/manifest/extensionpoint) 要素に含まれる子要素は 1 つでなければなりません。アドインがボタンを 1 つのみ追加する場合、子要素は [Control](/javascript/api/manifest/control) 要素になります。アドインがボタンを複数追加する場合、子要素は複数の `Control` 要素を含む [Group](/javascript/api/manifest/group) 要素になります。
- `Control` 要素に相当する `Menu` の種類はありません。
- [Supertip](/javascript/api/manifest/supertip) 要素は使用されません。
- アイコンの必須サイズが異なります。モバイル アドインは少なくとも 25x25、32x32 および 48x48 ピクセルのアイコンをサポートする必要があります。

## <a name="code-considerations"></a>コードに関する考慮事項

モバイル用のアドインの設計には、追加の考慮事項がいくつか導入されています。

### <a name="use-rest-instead-of-exchange-web-services"></a>Exchange Web サービスの代わりに REST を使用する

[Office.context.mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) メソッドは、Outlook Mobile ではサポートされていません。可能な場合には、アドインは優先的に Office.js API から情報を取得します。Office.js API によって表示されていない情報がアドインで必要な場合、[Outlook REST APIs](/outlook/rest/) を使用してユーザーのメールボックスにアクセスする必要があります。

メールボックス要件セット 1.5 には、REST API と互換性のあるアクセス トークンを要求できる新しいバージョンの [Office.context.mailbox.getCallbackTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) と、ユーザーの REST API エンドポイントを検索するために使用できる新しい [Office.context.mailbox.restUrl](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#properties) プロパティが導入されました。

### <a name="pinch-zoom"></a>ピンチによるズーム

既定で、ユーザーは "ピンチによるズーム" ジェスチャを使用して作業ウィンドウで拡大することができます。ご使用のシナリオでこれが該当しない場合は、HTML でピンチによるズームを無効にしてください。

### <a name="close-task-panes"></a>作業ウィンドウを閉じる

Outlook Mobile では、作業ウィンドウが画面全体を占めるので、既定ではユーザーが作業ウィンドウを閉じてメッセージに戻る必要があります。シナリオが完成したら、[Office.context.ui.closeContainer](/javascript/api/office/office.ui#office-office-ui-closecontainer-member(1)) メソッドを使用して作業ウィンドウを閉じることを検討してください。

### <a name="compose-mode-and-appointments"></a>作成モードと予定

現在、Outlook Mobile のアドインでは、メッセージを読み取るときにのみアクティブ化がサポートされています。 メッセージを作成するときや、予定を表示または作成するときには、アドインはアクティブ化されません。 ただし、次の 2 つの例外があります。

1. オンライン会議プロバイダー統合アドインは、予定開催者モードでアクティブ化できます。 この例外 (使用可能な API を含む) の詳細については、「 [オンライン会議プロバイダー用の Outlook モバイル アドインを作成する」](online-meeting.md#available-apis)を参照してください。
1. 予定のメモやその他の詳細を顧客関係管理 (CRM) またはメモ作成サービスに記録するアドインは、予定出席者モードでアクティブ化できます。 この例外 (使用可能な API を含む) の詳細については、「 [Outlook モバイル アドインの外部アプリケーションへの予定ノートのログ記録](mobile-log-appointments.md#available-apis)」を参照してください。

### <a name="unsupported-apis"></a>サポートされていない API

要件セット 1.6 以降で導入された API は、Outlook Mobile ではサポートされていません。 以前の要件セットの次の API もサポートされていません。

- [Office.context.officeTheme](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context#officetheme-officetheme)
- [Office.context.mailbox.ewsUrl](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#properties)
- [Office.context.mailbox.convertToEwsId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.convertToRestId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.displayAppointmentForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.displayMessageForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.displayNewAppointmentForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.item.dateTimeModified](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
- [Office.context.mailbox.item.displayReplyAllForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.displayReplyForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.getEntities](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.getEntitiesByType](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.getFilteredEntitiesByName](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.getRegexMatches](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.getRegexMatchesByName](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)

## <a name="see-also"></a>関連項目

[Exchange サーバーと Outlook クライアントでサポートされる要件セット](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients)