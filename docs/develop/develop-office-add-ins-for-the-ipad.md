---
title: iPad 用の Office アドインを開発する
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 77e67c361d227babebdd081ecdf308fc7469e507
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2018
ms.locfileid: "23944355"
---
# <a name="develop-office-add-ins-for-the-ipad"></a>iPad 用の Office アドインを開発する


次の表に、Office for iPad で実行する Office アドインを開発するときに実行するタスクの一覧を示します。


|**タスク**|**説明**|**リソース**|
|:-----|:-----|:-----|
|アドインを更新して、Office.js バージョン 1.1 をサポートします。|Office アドイン プロジェクトで使用する JavaScript ファイル (Office.js ファイルとアプリに固有の .js ファイル) とアドイン マニフェスト検証ファイルをバージョン 1.1 に更新します。|[JavaScript API for Office の変更点](https://docs.microsoft.com/javascript/office/what's-changed-in-the-javascript-api-for-office?view=office-js)|
|UI デザインのベスト プラクティスを適用します。|アドイン UI を iOS エクスペリエンスとシームレスに統合します。|[iOS の設計](https://developer.apple.com/library/ios/documentation/UserExperience/Conceptual/MobileHIG/)|
|アドイン デザインのベスト プラクティスを適用します。|アドインが明確な価値を提供し、魅力的であり、一貫して機能することを確認します。|[Office アドイン開発のベスト プラクティス](../concepts/add-in-development-best-practices.md)|
|タッチ用にアドインを最適化します。|マウスとキーボードに加え、タッチ入力に対して、UI が素早く応答するようにします。|[UX 設計原則を適用する](../concepts/add-in-development-best-practices.md#apply-ux-design-principles)|
|アドインを無料にします。|iPad 上の Office は、ユーザー数を拡大して、サービスを促進できるチャネルです。これらの新しいユーザーは、お客様になる可能性があります。|
  [検証ポリシー 10.8](https://docs.microsoft.com/office/dev/store/validation-policies#10-apps-and-add-ins-utilize-supported-capabilities)|
|アドインを商目的で使用しないようにします。|アドインは、アプリ内購入、試用版の提供、有料版へのアップセルを目的とする UI、またはユーザーが他のコンテンツやアプリやアドインを購入または取得できるすべてのオンライン ストアへのリンクと無縁である必要があります。またプライバシー ポリシーと使用条件のページにも、商用の UI または AppSource へのリンクがないことが必要です。|
  [検証ポリシー 3.4](https://docs.microsoft.com/office/dev/store/validation-policies#3-apps-and-add-ins-can-sell-additional-features-or-content-through-purchases-within-the-app-or-add-in)|
|アドインを AppSource に再送信します。|販売者ダッシュボードで、**[このアドインを iPad の Office アドイン カタログで利用できる状態にする]** チェック ボックスをオンにして、[Apple ID] ボックスに Apple 開発者 ID を入力します。[アプリケーション プロバイダー契約](https://sellerdashboard.microsoft.com/Assets/Content/Agreements/en-US/Office_Store_Seller_Agreement_20120927.htm)を確認して、契約を十分に理解します。|
  [AppSource と Office 内でソリューションを使用できるようにする](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store)|

他のプラットフォームで実行されている Office アプリケーション用にアドインをそのまま保持することができます。また、アドインが実行されているブラウザーとデバイスに基づく別の UI も提供できます。iPad 上でアドインが実行されているかどうかを検出するためには、次の API を使用できます。
- var isTouchEnabled = [Office.context.touchEnabled](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js#touchenabled)
- var allowCommerce = [Office.context.commerceAllowed](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js#commerceallowed)
    

## <a name="best-practices-for-developing-office-add-ins-for-ios-and-mac"></a>iOS および Mac 用 Office アドイン開発のベスト プラクティス

iOS 上で実行するアドインを開発するための次のベスト プラクティスを適用します。


-  **アドインの開発に Visual Studio を使用する。**
    
    アドインを Visual Studio で開発する場合、アドインを iPad または Mac にサイドロードする前に、Windows で動作する Office ホスト アプリケーションで、[そのコードのブレークポイントを設定してコードをデバッグ](../develop/create-and-debug-office-add-ins-in-visual-studio.md)できます。Office for iOS または Office for Mac で動作するアドインは Office for Windows で動作するアドインと同じ API をサポートするので、アドインのコードはどちらのプラットフォームでも同じように実行されるはずです。
    
-  **アドインのマニフェストまたはランタイム チェックを使用して API の要件を指定する。**
    
    アドインのマニフェストで API の要件を指定すると、Office はホスト アプリケーションがそれらの API メンバーをサポートするかどうかを調べます。API メンバーをホストで使用できる場合は、そのホスト アプリケーションでアドインを使用できます。または、ランタイム チェックを実行して、アドインで使用する前に、メソッドをホストで使用できるかどうかを確認することもできます。ランタイム チェックでは、ホストでアドインを常に使用できることが確認され、メソッドが使用可能な場合は追加機能が提供されます。詳細については、「[Office ホストと API 要件を指定する](specify-office-hosts-and-api-requirements.md)」を参照してください。
    
一般的なアドイン開発のベスト プラクティスについては、「[Office アドイン開発のベスト プラクティス](../concepts/add-in-development-best-practices.md)」を参照してください。


## <a name="see-also"></a>関連項目

- [iPad と Mac で Office アドインをサイドロードする](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)  
- [iPad と Mac で Office アドインをデバッグする](../testing/debug-office-add-ins-on-ipad-and-mac.md)
    
