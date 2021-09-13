---
title: Outlook Mobile の Outlook のアドイン
description: Outlookモバイル アドインは、すべてのビジネス アカウントと Microsoft 365.com アカウントOutlookサポートされています。
ms.date: 09/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 312e333e41dae5fa061540bec5018d4d2034b66b
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59149604"
---
# <a name="add-ins-for-outlook-mobile"></a>Outlook Mobile のアドイン

現時点で、アドインは他の Outlook エンドポイントで利用できるものと同じ API を使用して Outlook Mobile で動作します。Outlook 用のアドインを作成済みの場合、簡単に Outlook Mobile で動作するようにできます。

Outlookモバイル アドインは、すべてのビジネス アカウントと Microsoft 365.com アカウントOutlookサポートされています。 ただし、現在 Gmail アカウントではサポートを利用できません。

**Outlook on iOS の作業ウィンドウの例**

![iOS 上の作業ウィンドウのOutlookスクリーンショット。](../images/outlook-mobile-addin-taskpane.png)

<br/>

**Outlook on Android の作業ウィンドウの例**

![Android 上の作業ウィンドウのOutlookスクリーンショット。](../images/outlook-mobile-addin-taskpane-android.png)

> [!IMPORTANT]
> アドインは、モバイル ブラウザーの最新バージョンOutlook機能しません。 詳細については、「モバイル ブラウザー [Outlookアップグレード中」を参照してください](https://techcommunity.microsoft.com/t5/outlook-blog/outlook-on-your-mobile-browser-is-being-upgraded/ba-p/1125816)。

## <a name="whats-different-on-mobile"></a>モバイルにおける違い

- モバイル用の設計において、小さいサイズと迅速な操作性が課題となります。お客様に高品質のエクスペリエンスを提供するため、モバイル サポートを宣言するアドインに対して厳格な検証条件を定めています。AppSource で承認を得るには、この条件を満たす必要があります。
  - アドインは [UI ガイドライン](outlook-addin-design.md)に準拠 **していなければなりません**。
  - アドインのシナリオは、[モバイルに対して適切](#what-makes-a-good-scenario-for-mobile-add-ins)である **必要** があります。

- 一般に、現時点ではメッセージ読み取りモードだけがサポートされます。 つまり、マニフェストのモバイル セクションで宣言する必要がある唯一の `MobileMessageReadCommandSurface` [ExtensionPoint](../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface) です。 ただし、予定オーガナイザー モードは [、MobileOnlineMeetingCommandSurface](../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface)拡張ポイントを宣言するオンライン会議プロバイダー統合アドインでサポートされています。 このシナリオ[の詳細についてはOutlook](online-meeting.md)会議プロバイダーのモバイル アドインの作成に関する記事を参照してください。

- [makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) API はモバイルではサポートされていません。モバイル アプリは REST API を使用して、サーバーと通信します。アプリのバックエンドで Exchange サーバーと接続する必要がある場合、コールバック トークンを使用して REST API 呼び出しを行うことができます。詳しくは、「[Outlook アドインからの Outlook REST API の使用](use-rest-api.md)」をご覧ください。

- マニフェストで [MobileFormFactor](../reference/manifest/mobileformfactor.md) を使用してストアにアドインを送信する場合、iOS のアドインに関する当社の開発者補遺に同意し、確認のため Apple の開発者 ID を送信しなければなりません。

- 最後に、マニフェストで `MobileFormFactor` を宣言し、適切な種類の[コントロール](../reference/manifest/control.md)と[アイコンのサイズ](../reference/manifest/icon.md)を含める必要があります。

## <a name="what-makes-a-good-scenario-for-mobile-add-ins"></a>モバイル アドインに対して優れたシナリオにするには

電話での Outlook セッションの平均の長さは PC よりも短いことを忘れないでください。つまり、アドインを高速にする必要があります。さらに、シナリオでは、ユーザーの電子メール フローに出入りし、中断せずに続行できるようにする必要もあります。

Outlook Mobile に対して適切なシナリオの例を次に示します。

- アドインを使用すると、貴重な情報を Outlook に伝えることができるため、ユーザーは電子メールをトリアージし、適切に対応できます。例: ユーザーが顧客情報を確認し、適切な情報を共有するための CRM アドイン。

- アドインが、追跡システム、共同作業システム、または類似するシステムに情報を保存して、ユーザーの電子メール コンテンツに価値を追加します。例: ユーザーが電子メールを、プロジェクト進捗管理用にタスク項目に変換したり、サポート チーム用にヘルプ チケットに変換したりするアドイン。

**iOS で電子メール メッセージから Trello カードを作成するユーザーの操作の例**

![iOS 上のモバイル アドインOutlookユーザーの操作を示すアニメーション GIF。](../images/outlook-mobile-addin-interaction.gif)

<br/>

**Android で電子メール メッセージから Trello カードを作成するユーザーの操作の例**

![Android 上のモバイル アドインOutlook操作を示すアニメーション GIF。](../images/outlook-mobile-addin-interaction-android.gif)

## <a name="testing-your-add-ins-on-mobile"></a>モバイル上でのアドインのテスト

Outlook Mobile でアドインをテストするために、O365 や Outlook.com アカウントにアドインをサイドローディングできます。Outlook on the web で、設定ギアに移動し、[**統合の管理**] または [**アドインの管理**] を選択します。上部付近で、[**カスタム アドインを追加するには、ここをクリックします**] をクリックし、マニフェストをアップロードします。マニフェストの形式に `MobileFormFactor` が含まれていることを確認します。含まれていないと、読み込むことができません。

アドインが動作することを確認したら、携帯電話やタブレットなど、別のサイズの画面でテストします。コンストラストやフォント サイズ、色、さらには VoiceOver (iOS) または TalkBack (Android) などのスクリーン リーダーが使用できることなど、アクセシビリティのガイドラインに従っていることも確認してください。

モバイルでのトラブルシューティングは、使い慣らされたツールを使用していない可能性があるから、難しい場合があります。 ただし、iOS でのトラブルシューティングの 1 つの方法は、Fiddler を使用する方法です (iOS デバイスでの使用に関するこのチュートリアル [を参照してください](https://www.telerik.com/blogs/using-fiddler-with-apple-ios-devices))。

## <a name="next-steps"></a>次の手順

方法はこちら: 

- [モバイル サポートをアドインのマニフェストに追加する](add-mobile-support.md)。
- [アドインで優れたモバイル エクスペリエンスを設計する](outlook-addin-design.md)。
- アドインから[アクセス トークンを取得し、Outlook REST API を呼び出す](use-rest-api.md)。
