---
title: Office アドインの認証設計ガイドライン
ms.date: 02/09/2021
description: Office アドインでサインオン ページまたはサインアップ ページを視覚的に設計する方法について説明します。
ms.localizationpriority: medium
ms.openlocfilehash: 45d11d509585a199135889273e6f9a96ce98e691
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810261"
---
# <a name="authentication-patterns"></a>認証パターン

アドインの機能にユーザーがアクセスするには、サインインまたはサインアップする必要があります。 認証時の典型的なインターフェイス コントロールには、ユーザー名とパスワードの入力ボックスやサードパーティの資格情報フローを開始するボタンがあります。 ユーザーにアドインの使用を開始してもらうには、簡単で効率的に認証を導入することが重要な最初の一歩となります。

## <a name="best-practices"></a>ベスト プラクティス

|するべきこと|してはいけないこと|
|:----|:----|
|サインインの前に、アドインの価値について説明し、アカウントを要求せずにこの機能を実際に使用します。 |アドインの価値と長所を理解せずにサインインすることをユーザーに期待します。|
|各画面に目立つ第 1 のボタンを配置し、ユーザーに認証フローを段階的に説明します。 |ボタンや行動喚起が競合する第 2 のタスクや第 3 のタスクに注意を向けさせます。|
|"サインイン" や "アカウント作成" など、特定のタスクを説明するわかりやすいボタン ラベルを使用します。 |認証フローでユーザーを誘導するとき、"送信" や "開始" のようなあいまいなボタン ラベルを使用します。|
|ダイアログを使用し、ユーザーの注意を認証フォームに向けさせます。 |最初の実行エクスペリエンスと認証フォームで作業ウィンドウをあふれさせます。|
|入力ボックスの自動フォーカスなど、フローの中に小さな効率性を見つけます。 |クリックしてフォーム フィールドに入るようにユーザーに要求するなど、操作に不要な手順を追加します。|
|ユーザーがサインアウトして再認証する方法を提供します。 |ID を切り替える際、アンインストールをユーザーに強制します。|

## <a name="authentication-flow"></a>認証フロー

1. 初回実行プレースマット - アドインの最初の実行エクスペリエンス内にわかりやすい行動喚起としてサインイン ボタンを配置します。

    ![Office アプリケーションのアドイン作業ウィンドウを示すスクリーンショット。](../images/add-in-fre-value-placemat.png)

1. ID プロバイダーの選択肢ダイアログ - ID プロバイダーのわかりやすい一覧を表示します。該当する場合、ユーザー名やパスワードのフォームも含めます。 認証ダイアログが開いているとき、アドイン UI はブロックされることがあります。

    ![Office アプリケーションの [ID プロバイダーの選択肢] ダイアログを示すスクリーンショット。](../images/add-in-auth-choices-dialog.png)

1. ID プロバイダーのサインイン - ID プロバイダーによって独自の UI が提供されます。 Microsoft Azure Active Directoryを使用すると、サインインページとアクセス パネル ページをカスタマイズして、サービスとの一貫性のある外観を実現できます。 [詳細については、こちらをご覧ください](/azure/active-directory/fundamentals/customize-branding)。

    ![Office アプリケーションの [ID プロバイダーのサインイン] ダイアログを示すスクリーンショット。](../images/add-in-auth-identity-sign-in.png)

1. 進捗状況 - 設定や UI の読み込みの進行状況を示します。

    ![Office アプリケーションの進行状況インジケーターを含むダイアログを示すスクリーンショット。](../images/add-in-auth-modal-interstitial.png)

> [!NOTE]
> Microsoft の ID サービスを使用すると、商標付きのサインイン ボタンを使用できます。このボタンは淡色テーマまたは濃色テーマにカスタマイズできます。 詳細情報。

## <a name="single-sign-on-authentication-flow"></a>単一Sign-On認証フロー

> [!NOTE]
> シングル サインオン API は現在、Word、Excel、Outlook、PowerPoint でサポートされています。 シングル サインオンのサポートの詳細については、「 [IdentityAPI 要件セット](/javascript/api/requirement-sets/common/identity-api-requirement-sets)」を参照してください。 Outlook アドインで作業している場合は、Microsoft 365 テナントの先進認証が有効になっていることを確認してください。 この方法の詳細については、「[Exchange Online: テナントの先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」を参照してください。

よりスムーズなエンド ユーザー エクスペリエンスを実現するには、シングル サインオンを使用します。 Office 内のユーザーの ID (Microsoft アカウントまたは Microsoft 365 ID) は、アドインへのサインインに使用されます。 そのため、ユーザーは 1 回だけサインインします。 お客様は途中で止められることなく、簡単に利用を開始できます。

1. アドインがインストールされると、次のような同意ウィンドウがユーザーに表示されます。

    ![アドインのインストール時に Office アプリケーションの同意ウィンドウを示すスクリーンショット。](../images/add-in-auth-SSO-consent-dialog.png)

    > [!NOTE]
    > この同意ウィンドウに含まれるロゴ、文字列、アクセス許可の範囲については、アドインの発行元が制御します。 UI は Microsoft が事前に構成したものです。

1. アドインはユーザーが同意した後で読み込まれます。 ユーザーがカスタマイズした情報が必要であれば、それを抽出し、表示できます。

    ![リボンにアドイン ボタンが表示された Office アプリケーションを示すスクリーンショット。](../images/add-in-ribbon.png)

## <a name="see-also"></a>関連項目

- [SSO アドインの開発の](../develop/sso-in-office-add-ins.md)詳細
