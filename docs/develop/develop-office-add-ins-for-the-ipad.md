---
title: iPad アドインの特別な要件
description: iPad で実行される Office アドインを作成するための要件について説明します。
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: cc75cc75daec756efcb066f3e3a77f865672e501
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889304"
---
# <a name="special-requirements-for-add-ins-on-the-ipad"></a>iPad アドインの特別な要件

アドインで iPad でサポートされている Office API のみを使用している場合は、お客様が iPad にインストールできます。 (詳細については、「[Office アプリケーションと API 要件の指定](specify-office-hosts-and-api-requirements.md)」を参照してください)。*[アドインが AppSource](https://appsource.microsoft.com) を通じて販売される場合* は、[すべての Office アドインに適用されるベスト プラクティス](../concepts/add-in-development-best-practices.md)に加えて、iPad にインストールできるアドインに対して従う必要があるいくつかのプラクティスがあります。

次の表に、実行するタスクを示します。

> [!NOTE]
> Outlook Mobile で適切に動作する Outlook アドインの設計については、「 [Outlook Mobile 用アドイン](../outlook/outlook-mobile-addins.md)」を参照してください。

|タスク|説明|リソース|
|:-----|:-----|:-----|
|アドインを更新して、Office.js バージョン 1.1 をサポートします。|Office アドイン プロジェクトで使用する JavaScript ファイル (Office.js ファイルとアプリに固有の .js ファイル) とアドイン マニフェスト検証ファイルをバージョン 1.1 に更新します。|[API とマニフェストのバージョンを更新する](update-your-javascript-api-for-office-and-manifest-schema-version.md)|
|iOS デザインのベスト プラクティスを適用します。|アドイン UI を iOS エクスペリエンスとシームレスに統合します。| 以下の注を参照してください。 |
|タッチ用にアドインを最適化します。|マウスとキーボードに加え、タッチ入力に対して、UI が素早く応答するようにします。|[UX 設計原則を適用する](../concepts/add-in-development-best-practices.md#apply-ux-design-principles)|
|アドインを無料にします。|iPad 上の Office は、ユーザー数を拡大して、サービスを促進できるチャネルです。これらの新しいユーザーは、お客様になる可能性があります。|[認定ポリシー 1120.2](/legal/marketplace/certification-policies#11202-acquisition-pricing-and-terms)|
|iPad でアドインコマースを無料にします。|iPad で実行されている場合、アドインは、アプリ内購入、試用版オファー、非無料版へのアップセルを目的とした UI、またはユーザーが他のコンテンツ、アプリ、またはアドインを購入または取得できるオンライン ストアへのリンクが含まれなければなりません。プライバシー ポリシーと使用条件のページには、コマース UI または AppSource のリンクも含めなければなりません。|[認定ポリシー 1100.3](/legal/marketplace/certification-policies#11003-selling-additional-features)<br><br>アドインは、他のプラットフォームで引き続きコマースを行うことができます。 これを行うには、 [Office.context.commerceAllowed](/javascript/api/office/office.context#office-office-context-commerceallowed-member) プロパティをテストし、返されたときにすべてのコマースを抑制します `false`。|
|アドインを AppSource に送信します。|パートナー センターの [ **製品のセットアップ** ] ページで、[ **iOS と Android で製品を使用できるようにする (該当する場合)]** チェック ボックスをオンにし、アカウント設定で Apple 開発者 ID を入力します。 [アプリケーション プロバイダー契約](https://go.microsoft.com/fwlink/?linkid=715691)を確認して、条件を理解していることを確認します。|[AppSource と Office 内でソリューションを使用できるようにする](/office/dev/store/submit-to-appsource-via-partner-center)|

> [!NOTE]
> アドインは、実行されているデバイスに基づいて代替 UI を提供できます。 アドインが iPad で実行されているかどうかを検出するには、次の API を使用します。
>
> - const isTouchEnabled = [Office.context.touchEnabled](/javascript/api/office/office.context#office-office-context-touchenabled-member)
> - const allowCommerce = [Office.context.commerceAllowed](/javascript/api/office/office.context#office-office-context-commerceallowed-member)
>
> iPad では、戻り`true`値`touchEnabled`と`commerceAllowed`戻り値を返します`false`。
>
> iPad の最適な UI デザインプラクティスについては、「 [iOS 用の設計](https://developer.apple.com/library/ios/documentation/UserExperience/Conceptual/MobileHIG/)」を参照してください。

## <a name="best-practices-for-developing-office-add-ins-that-can-run-on-ipad"></a>iPad で実行できる Office アドインを開発するためのベスト プラクティス

iPad で実行されるアドインを開発するための次のベスト プラクティスを適用します。

- **Windows または Mac でアドインを開発してデバッグし、iPad にサイドロードします。**

    アドインは iPad で直接開発することはできませんが、Windows または Mac コンピューターで開発してデバッグし、テスト用に iPad にサイドロードできます。 Office on iOS または Mac で実行されるアドインは、Windows 上の Office で実行されているアドインと同じ API をサポートしているため、アドインのコードはこれらのプラットフォームでも同じように実行される必要があります。 詳細については、「 [テスト用に iPad で Office アドインをテストしてデバッグ](../testing/test-debug-office-add-ins.md) し、 [Office アドインをサイドロードする」を参照](../testing/sideload-an-office-add-in-on-ipad.md)してください。

- **アドインのマニフェストまたはランタイム チェックを使用して API の要件を指定する。**

    アドインのマニフェストで API 要件を指定すると、Office クライアント アプリケーションがそれらの API メンバーをサポートしているかどうかを Office が判断します。 API メンバーがアプリケーションで使用できる場合は、アドインを使用できます。 または、ランタイム チェックを実行して、アドインでメソッドを使用する前に、アプリケーションでメソッドを使用できるかどうかを判断することもできます。 ランタイム チェックでは、アドインがアプリケーションで常に使用できることを確認し、メソッドが使用可能な場合は追加機能を提供します。 詳細については、「 [Office アプリケーションと API 要件の指定](specify-office-hosts-and-api-requirements.md)」を参照してください。
