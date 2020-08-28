---
title: iPad 用の Office アドインを開発する
description: IPad で実行する Office アドインを作成するための概要とベストプラクティスについて説明します。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 6738cc559cc07f747e075c17419b70558dec3c66
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292786"
---
# <a name="develop-office-add-ins-for-the-ipad"></a>iPad 用の Office アドインを開発する


次の表に、office for iPad で実行する Office アドインを開発するために実行するタスクの一覧を示します。


|**タスク**|**説明**|**リソース**|
|:-----|:-----|:-----|
|アドインを更新して、Office.js バージョン 1.1 をサポートします。|Office アドイン プロジェクトで使用する JavaScript ファイル (Office.js ファイルとアプリに固有の .js ファイル) とアドイン マニフェスト検証ファイルをバージョン 1.1 に更新します。|[API とマニフェストのバージョンを更新する](update-your-javascript-api-for-office-and-manifest-schema-version.md)|
|UI デザインのベスト プラクティスを適用します。|アドイン UI を iOS エクスペリエンスとシームレスに統合します。|[iOS の設計](https://developer.apple.com/library/ios/documentation/UserExperience/Conceptual/MobileHIG/)|
|アドイン デザインのベスト プラクティスを適用します。|アドインが明確な価値を提供し、魅力的であり、一貫して機能することを確認します。|[Office アドイン開発のベスト プラクティス](../concepts/add-in-development-best-practices.md)|
|タッチ用にアドインを最適化します。|マウスとキーボードに加え、タッチ入力に対して、UI が素早く応答するようにします。|[UX 設計原則を適用する](../concepts/add-in-development-best-practices.md#apply-ux-design-principles)|
|アドインを無料にします。|iPad 上の Office は、ユーザー数を拡大して、サービスを促進できるチャネルです。これらの新しいユーザーは、お客様になる可能性があります。|[証明ポリシー1120.2](/legal/marketplace/certification-policies#11202-acquisition-pricing-and-terms)|
|アドインを商目的で使用しないようにします。|アドインは、アプリ内購入、試用版の提供、有料版へのアップセルを目的とする UI、またはユーザーが他のコンテンツやアプリやアドインを購入または取得できるすべてのオンライン ストアへのリンクと無縁である必要があります。またプライバシー ポリシーと使用条件のページにも、商用の UI または AppSource へのリンクがないことが必要です。|[証明ポリシー1100.3](/legal/marketplace/certification-policies#11003-selling-additional-features)|
|アドインを AppSource に再送信します。|[パートナーセンター] の [ **製品のセットアップ** ] ページで、[ **iOS および Android で製品を利用できるようにする (該当する場合)** ] チェックボックスをオンにして、[アカウント設定] に APPLE の開発者 ID を入力します。 [アプリケーションプロバイダアグリーメント](https://go.microsoft.com/fwlink/?linkid=715691)を確認して、用語を理解していることを確認してください。|[AppSource と Office 内でソリューションを使用できるようにする](/office/dev/store/submit-to-appsource-via-partner-center)|

他のプラットフォームで実行されている Office アプリケーション用にアドインをそのまま保持することができます。また、アドインが実行されているブラウザーとデバイスに基づく別の UI も提供できます。iPad 上でアドインが実行されているかどうかを検出するためには、次の API を使用できます。
- var isTouchEnabled = [Office.context.touchEnabled](/javascript/api/office/office.context#touchenabled)
- var allowCommerce = [Office.context.commerceAllowed](/javascript/api/office/office.context#commerceallowed)


## <a name="best-practices-for-developing-office-add-ins-for-ios-and-mac"></a>iOS および Mac 用 Office アドイン開発のベスト プラクティス

iOS 上で実行するアドインを開発するための次のベスト プラクティスを適用します。


-  **アドインの開発に Visual Studio を使用する。**

    Visual Studio を使用してアドインを開発する場合は、iPad または Mac でアドインをサイドロードする前に、Windows 上で実行されている Office クライアントアプリケーションで [ブレークポイントを設定し、コードをデバッグ](../develop/debug-office-add-ins-in-visual-studio.md) することができます。 IOS または Mac 上の Office で実行するアドインは、Windows の Office で実行されるアドインと同じ Api をサポートしているため、アドインのコードは両方のプラットフォームで同じように実行する必要があります。

-  **アドインのマニフェストまたはランタイム チェックを使用して API の要件を指定する。**

    アドインのマニフェストで API 要件を指定すると、office は Office クライアントアプリケーションがそれらの API メンバーをサポートしているかどうかを判断します。 API メンバーがアプリケーションで使用可能な場合は、アドインが使用できるようになります。 または、ランタイムチェックを実行して、アドインで使用する前に、そのメソッドがアプリケーションで使用できるかどうかを確認することもできます。 ランタイムチェックアドインが常にアプリケーションで使用できることを確認し、メソッドが使用可能な場合は追加機能を提供します。 詳細については、「 [Office アプリケーションと API 要件を指定する](specify-office-hosts-and-api-requirements.md)」を参照してください。

一般的なアドイン開発のベスト プラクティスについては、「[Office アドイン開発のベスト プラクティス](../concepts/add-in-development-best-practices.md)」を参照してください。


## <a name="see-also"></a>関連項目

- [iPad または Mac で Office アドインをサイドロードする](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)  
- [iPad と Mac で Office アドインをデバッグする](../testing/debug-office-add-ins-on-ipad-and-mac.md)
