
# <a name="develop-office-add-ins-for-the-ipad"></a>iPad 用の Office アドインを開発する


次の表に、Office for iPad で実行する Office アドインを開発するときに実行するタスクの一覧を示します。


|**タスク**|**説明**|**リソース**|
|:-----|:-----|:-----|
|アドインを更新して、Office.js バージョン 1.1 をサポートします。|Office アドイン プロジェクトで使用する JavaScript ファイル (Office.js ファイルとアプリに固有の .js ファイル) とアドイン マニフェスト検証ファイルをバージョン 1.1 に更新します。|[JavaScript API for Office の変更点](https://dev.office.com/reference/add-ins/what's-changed-in-the-javascript-api-for-office)|
|UI デザインのベスト プラクティスを適用します。|アドイン UI を iOS エクスペリエンスとシームレスに統合します。|[iOS の設計](https://developer.apple.com/library/ios/documentation/UserExperience/Conceptual/MobileHIG/)|
|アドイン デザインのベスト プラクティスを適用します。|アドインが明確な価値を提供し、魅力的であり、一貫して機能することを確認します。|[Office アドイン開発のベスト プラクティス](../../docs/overview/add-in-development-best-practices.md)|
|タッチ用にアドインを最適化します。|マウスとキーボードに加え、タッチ入力に対して、UI が素早く応答するようにします。|[UX 設計原則を適用する](https://msdn.microsoft.com/EN-US/library/mt590883.aspx#Anchor_3)|
|アドインを無料にします。|iPad 上の Office は、ユーザー数を拡大して、サービスを促進できるチャネルです。これらの新しいユーザーは、お客様になる可能性があります。|[検証ポリシー 10.8](http://msdn.microsoft.com/library/cd90836a-523e-42f5-ab02-5123cdf9fefe%28Office.15%29.aspx)|
|アドインを商目的で使用しないようにします。|アドインは、アプリ内購入、試用版の提供、有料版へのアップセルを目的とする UI、またはユーザーが他のコンテンツやアプリやアドインを購入または取得できるすべてのオンライン ストアへのリンクと無縁である必要があります。またプライバシー ポリシーと使用条件のページにも、商用の UI またはストアへのリンクがないことが必要です。|[検証ポリシー 3.4](http://msdn.microsoft.com/library/cd90836a-523e-42f5-ab02-5123cdf9fefe%28Office.15%29.aspx)|
|アドインを Office ストアに再送信します。|販売者ダッシュボードで、**[このアドインを iPad の Office アドイン カタログで利用できる状態にする]** チェック ボックスをオンにして、[Apple ID] ボックスに Apple 開発者 ID を入力します。[Office ストア アプリケーション プロバイダー契約](https://sellerdashboard.microsoft.com/Assets/Content/Agreements/en-US/Office_Store_Seller_Agreement_20120927.htm)を確認して、契約を十分に理解します。|[Office ストアに Office アドインと SharePoint アドインおよび Office 365 Web アプリを提出する](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx)|

他のプラットフォームで実行されている Office アプリケーション用にアドインをそのまま保持することができます。また、アドインが実行されているブラウザーとデバイスに基づく別の UI も提供できます。iPad 上でアドインが実行されているかどうかを検出するためには、次の API を使用できます。<ul><li>var isTouchEnabled = [Office.context.touchEnabled](http://dev.office.com/reference/add-ins/shared/office.context.touchenabled)</li><li>var allowCommerce = [Office.context.commerceAllowed](http://dev.office.com/reference/add-ins/shared/office.context.commerceallowed)</li></ul>
    

## <a name="best-practices-for-developing-office-add-ins-for-ios-and-mac"></a>iOS および Mac 用 Office アドイン開発のベスト プラクティス

iOS 上で実行するアドインを開発するための次のベスト プラクティスを適用します。


-  **アドインの開発に Visual Studio を使用する。**
    
    アドインを Visual Studio で開発する場合、アドインを iPad または Mac にサイドロードする前に、Windows で動作する Office ホスト アプリケーションで、[そのコードのブレークポイントを設定してコードをデバッグ](../get-started/create-and-debug-office-add-ins-in-visual-studio.md#Test)できます。Office for iOS または Office for Mac で動作するアドインは Office for Windows で動作するアドインと同じ API をサポートするので、アドインのコードはどちらのプラットフォームでも同じように実行されるはずです。
    
-  **アドインのマニフェストまたはランタイム チェックを使用して API の要件を指定する。**
    
    アドインのマニフェストで API の要件を指定すると、Office はホスト アプリケーションがそれらの API メンバーをサポートするかどうかを調べます。API メンバーをホストで使用できる場合は、そのホスト アプリケーションでアドインを使用できます。または、ランタイム チェックを実行して、アドインで使用する前に、メソッドをホストで使用できるかどうかを確認することもできます。ランタイム チェックでは、ホストでアドインを常に使用できることが確認され、メソッドが使用可能な場合は追加機能が提供されます。詳細については、「[Office ホストと API 要件を指定する](../../docs/overview/specify-office-hosts-and-api-requirements.md)」を参照してください。
    
一般的なアドイン開発のベスト プラクティスについては、「[Office アドイン開発のベスト プラクティス](../../docs/overview/add-in-development-best-practices.md)」を参照してください。


## <a name="additional-resources"></a>その他のリソース
<a name="bk_addresources"> </a>


- [iPad または Mac で Office アドインをサイドロードする](../../docs/testing/sideload-an-office-add-in-on-ipad-and-mac.md)
    
- [iPad と Mac で Office アドインをデバッグする](../../docs/testing/debug-office-add-ins-on-ipad-and-mac.md)
    
