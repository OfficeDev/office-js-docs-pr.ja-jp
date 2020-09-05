---
title: IPad でのアドインの特別な要件
description: IPad で実行する Office アドインを作成するためのいくつかの要件について説明します。
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: 25ac5767db3301352e1921411af833957c4644d0
ms.sourcegitcommit: 10463841a977e9b8415362a3ae91b0ae5eebbf89
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/04/2020
ms.locfileid: "47399572"
---
# <a name="special-requirements-for-add-ins-on-the-ipad"></a>IPad でのアドインの特別な要件

アドインが iPad でサポートされている Office Api のみを使用している場合、お客様は Ipad にインストールできます。 (詳細については[、「Office アプリケーションと API 要件を指定する](specify-office-hosts-and-api-requirements.md)」を参照してください)。*アドインが[appsource](https://appsource.microsoft.com)を使用して市場に配置*される場合は、[すべての Office アドインに適用されるベストプラクティス](../concepts/add-in-development-best-practices.md)に加えて、ipad にインストールできるアドインに従う必要があるいくつかのプラクティスがあります。

次の表に、実行するタスクを示します。

> [!NOTE]
> Outlook Mobile で正常に動作する Outlook アドインを設計する方法については、「 [Outlook mobile 用のアドイン](../outlook/outlook-mobile-addins.md)」を参照してください。

|タスク|説明|リソース|
|:-----|:-----|:-----|
|アドインを更新して、Office.js バージョン 1.1 をサポートします。|Office アドイン プロジェクトで使用する JavaScript ファイル (Office.js ファイルとアプリに固有の .js ファイル) とアドイン マニフェスト検証ファイルをバージョン 1.1 に更新します。|[API とマニフェストのバージョンを更新する](update-your-javascript-api-for-office-and-manifest-schema-version.md)|
|IOS 設計のベストプラクティスを適用します。|アドイン UI を iOS エクスペリエンスとシームレスに統合します。| 以下のメモを参照してください。 |
|タッチ用にアドインを最適化します。|マウスとキーボードに加え、タッチ入力に対して、UI が素早く応答するようにします。|[UX 設計原則を適用する](../concepts/add-in-development-best-practices.md#apply-ux-design-principles)|
|アドインを無料にします。|iPad 上の Office は、ユーザー数を拡大して、サービスを促進できるチャネルです。これらの新しいユーザーは、お客様になる可能性があります。|[証明ポリシー1120.2](/legal/marketplace/certification-policies#11202-acquisition-pricing-and-terms)|
|アドインを iPad 上で無料にしてください。|IPad で実行している場合、アドインはアプリ内購入、試用版、無料版へのアップセルを対象とする UI、またはユーザーが他のコンテンツ、アプリ、またはアドインを購入または取得できる任意のオンラインストアへのリンクを無料で使用する必要があります。プライバシーポリシーと使用条件ページも、commerce UI または AppSource リンクからすべて解放されている必要があります。|[証明ポリシー1100.3](/legal/marketplace/certification-policies#11003-selling-additional-features)<br><br>アドインでは、他のプラットフォームで commerce を引き続き利用できます。 これを行うには、 [commerceAllowed](/javascript/api/office/office.context#commerceallowed) プロパティをテストし、返されるすべての商取引を抑制し `false` ます。|
|アドインを AppSource に提出します。|[パートナーセンター] の [ **製品のセットアップ** ] ページで、[ **iOS および Android で製品を利用できるようにする (該当する場合)** ] チェックボックスをオンにして、[アカウント設定] に APPLE の開発者 ID を入力します。 [アプリケーションプロバイダアグリーメント](https://go.microsoft.com/fwlink/?linkid=715691)を確認して、用語を理解していることを確認してください。|[AppSource と Office 内でソリューションを使用できるようにする](/office/dev/store/submit-to-appsource-via-partner-center)|

> [!NOTE]
> アドインは、を実行しているデバイスに基づいて、代替 UI を提供することができます。 アドインが iPad 上で実行されているかどうかを検出するには、次の Api を使用できます。
>
> - var isTouchEnabled = [Office.context.touchEnabled](/javascript/api/office/office.context#touchenabled)
> - var allowCommerce = [Office.context.commerceAllowed](/javascript/api/office/office.context#commerceallowed)
>
> IPad では、 `touchEnabled` 戻り、を `true` `commerceAllowed` 返し `false` ます。
>
> IPad の最適な UI 設計手法の詳細については、「 [iOS 向けの設計](https://developer.apple.com/library/ios/documentation/UserExperience/Conceptual/MobileHIG/)」を参照してください。

## <a name="best-practices-for-developing-office-add-ins-that-can-run-on-ipad"></a>IPad で実行できる Office アドインを開発するためのベストプラクティス

IPad 上で実行するアドインを開発するための次のベストプラクティスを適用します。

-  **Windows または Mac でアドインを開発およびデバッグし、iPad にサイドロードます。**

    IPad でアドインを直接開発することはできませんが、Windows または Mac コンピューターで開発およびデバッグしてテスト用にサイドロードことができます。 IOS または Mac で Office で実行するアドインは、Windows の Office で実行されるアドインと同じ Api をサポートしているため、アドインのコードをこれらのプラットフォームで同じように実行する必要があります。 詳細については、「 [テストとデバッグ」の「Office アドインのテストとデバッグ](../testing/test-debug-office-add-ins.md) 」、および「 [テスト用に iPad と Mac で Office アドインをサイドロードする](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)」を参照してください。

-  **アドインのマニフェストまたはランタイム チェックを使用して API の要件を指定する。**

    アドインのマニフェストで API 要件を指定すると、office は Office クライアントアプリケーションがそれらの API メンバーをサポートしているかどうかを判断します。 API メンバーがアプリケーションで使用可能な場合は、アドインが使用できるようになります。 または、ランタイムチェックを実行して、アドインで使用する前に、そのメソッドがアプリケーションで使用できるかどうかを確認することもできます。 ランタイムチェックアドインが常にアプリケーションで使用できることを確認し、メソッドが使用可能な場合は追加機能を提供します。 詳細については、「 [Office アプリケーションと API 要件を指定する](specify-office-hosts-and-api-requirements.md)」を参照してください。
