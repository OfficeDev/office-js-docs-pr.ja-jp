---
title: Outlook コンテキスト アドインのアクティブ化のトラブルシューティング
description: アドインが想定どおりにアクティブ化されない可能性がある理由。
ms.date: 08/09/2022
ms.localizationpriority: medium
ms.openlocfilehash: c0034eccc1143e3af9867702cdf7cefa6f6a8c53
ms.sourcegitcommit: 57258dd38507f791bbb39cbb01d6bbd5a9d226b9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/12/2022
ms.locfileid: "67318887"
---
# <a name="troubleshoot-outlook-add-in-activation"></a>Outlook アドインのアクティブ化のトラブルシューティング

Outlook コンテキスト アドインのアクティブ化は、アドイン マニフェストのアクティブ化規則に基づいています。 現在選択されている項目の条件がアドインのアクティブ化規則を満たしている場合、アプリケーションは Outlook UI でアドイン ボタンをアクティブ化して表示します (作成アドインのアドイン選択ウィンドウ、読み取りアドインのアドイン バー)。 しかし、アドインが想定どおりにアクティブ化されない場合、考えられる理由を探るために次のような点を調べる必要があります。

## <a name="is-user-mailbox-on-a-version-of-exchange-server-that-is-at-least-exchange-2013"></a>ユーザーのメールボックスが、Exchange 2013 以降のバージョンの Exchange Server 上にあるか?

まず、テストしているユーザーの電子メール アカウントが、Exchange 2013 以降のバージョンの Exchange Server 上にあることを確認します。Exchange 2013 より後にリリースされた特定の機能を使用する場合は、ユーザーのアカウントが Exchange の適切なバージョン上にあることを確認してください。

次のいずれかの方法を使用して、Exchange 2013 のバージョンを確認できます。

- Exchange Server 管理者に確認します。

- スクリプト デバッガー (たとえば、Internet Explorer に付属する JScript デバッガーなど) で Outlook on the web またはモバイル デバイス上のアドインをテストしている場合は、スクリプトの読み込み元を指定する **script** タグの **src** 属性を探します。このパスには、**owa/15.0.516.x/owa2/...** という部分文字列があります。この中の **15.0.516.x** が Exchange Server のバージョン (**15.0.516.2** など) を表します。

- あるいは、[Office.context.mailbox.diagnostics.hostVersion](/javascript/api/outlook/office.diagnostics#outlook-office-diagnostics-hostversion-member) プロパティを使用してバージョンを確認することもできます。Outlook on the web およびモバイル デバイス上で、このプロパティは Exchange Server のバージョンを返します。

- Outlook でアドインをテストできる場合は、Outlook オブジェクト モデルと Visual Basic エディターを使用する次の簡単なデバッグ手法を使用できます。

    1. 最初に、Outlook でマクロが有効になっていることを確認します。**[ファイル]**、**[オプション]**、**[セキュリティ センター]**、**[セキュリティ センターの設定]**、**[マクロの設定]** の順に選択します。セキュリティ センターで、**[すべてのマクロの通知]** が選択されていることを確認します。Outlook の起動時に **[マクロを有効にする]** も選択している必要があります。

    1. リボンの **[開発]** タブで **[Visual Basic]** を選択します。

       > [!NOTE]
       > **[開発]** タブが表示されない場合には、「[方法:[開発] タブをリボンに表示する](/visualstudio/vsto/how-to-show-the-developer-tab-on-the-ribbon)」を参照して、有効にしてください。

    1. Visual Basic エディターで、**[表示]**、**[イミディエイト ウィンドウ]** を選択します。

    1. イミディエイト ウィンドウに次のように入力し、Exchange Server のバージョンを表示します。戻される値のメジャー バージョンは、15 以上である必要があります。

       - ユーザーのプロファイルに Exchange アカウントが 1 つだけある場合:

       ```vb
        ?Session.ExchangeMailboxServerVersion
       ```

       - 同じユーザー プロファイルに複数の Exchange アカウントがある場合 (`emailAddress` はユーザーのプライマリ STMP アドレスを含む文字列を表します):

       ```vb
        ?Session.Accounts.Item(emailAddress).ExchangeMailboxServerVersion
       ```

## <a name="is-the-add-in-disabled"></a>アドインが無効化されていないか?

いずれかの Outlook リッチ クライアントで、パフォーマンス上の理由によりアドインが無効化されている可能性があります。たとえば、CPU コア使用率やメモリ使用量のしきい値、クラッシュ許容度、およびアドインに対するすべての正規表現の処理時間が超過した場合などです。このようなことが起きると、Outlook リッチ クライアントは、アドインを無効化していることを示す通知を表示します。

> [!NOTE]
> リソース使用量を監視するのは Outlook リッチ クライアントだけですが、Outlook リッチ クライアントでアドインを無効化すると、Outlook on the web とモバイル デバイスでもアドインが無効化されます。

アドインが無効になっているかどうかを確認するには、次のいずれかの方法を使用します。

- Outlook on the webで、電子メール アカウントに直接サインインし、リボンから [**アドインの取得**] を選択します。

- Outlook on Windows で、リボンから **[その他のアプリ** ] を選択し、[ **アドインの取得**] を選択します。

- Outlook on Mac で、リボンから省略記号ボタン (`...`) を選択し、[ **アドインの取得**] を選択します。

## <a name="does-the-tested-item-support-outlook-add-ins-is-the-selected-item-delivered-by-a-version-of-exchange-server-that-is-at-least-exchange-2013"></a>テストするアイテムが Outlook アドインをサポートしているか? 選択されたアイテムが Exchange 2013 以降のバージョンの Exchange Server で配信されているか?

Outlook アドインが閲覧アドインであり、ユーザーがメッセージ (メール メッセージ、会議出席依頼、返信、キャンセルなど) や予定を表示するときにアクティブ化されるものである場合、これらのアイテムが通常はアドインをサポートしているとしても、選択しているアイテムが次のいずれかの場合は例外があります。 選択したアイテムが、 [Outlook アドインがアクティブ化されない一覧](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins)の項目の 1 つであるかどうかを確認します。

また、予定は常にリッチ テキスト形式で保存されるため、**BodyAsHTML** の **PropertyName** 値を指定する [ItemHasRegularExpressionMatch](/javascript/api/manifest/rule#itemhasregularexpressionmatch-rule) ルールでは、プレーンテキストまたはリッチ テキスト形式で保存された予定またはメッセージに対してアドインがアクティブ化されません。

メール アイテムが上記の種類の 1 つではない場合でも、少なくとも Exchange 2013 のバージョンのExchange Serverによってアイテムが配信されなかった場合、送信者の SMTP アドレスなどの既知のエンティティとプロパティはアイテムで識別されません。 これらのエンティティまたはプロパティに依存するすべてのアクティブ化ルールは満たされておらず、アドインはアクティブ化されません。

Windows 以外の Outlook on client で、ユーザーがメッセージまたは会議出席依頼を作成しているときにアドインがアクティブ化される場合は、アイテムが Information Rights Management (IRM) によって保護されていないことを確認します。

[!INCLUDE [outlook-irm-add-in-activation](../includes/outlook-irm-add-in-activation.md)]

## <a name="is-the-add-in-manifest-installed-properly-and-does-outlook-have-a-cached-copy"></a>アドイン マニフェストが適切にインストールされているか? また Outlook にキャッシュ コピーがあるか?

このシナリオは、Windows 上の Outlook にのみ適用されます。 通常、メールボックスに Outlook アドインをインストールすると、Exchange Server は、アドイン マニフェストを指定の場所からその Exchange Server 上のメールボックスにコピーします。 Outlook が起動するたびに、そのメールボックスにインストールされているすべてのマニフェストが、次の場所にある一時的なキャッシュに読み取られます。

```text
%LocalAppData%\Microsoft\Office\16.0\WEF
```

たとえば、ユーザー John の場合、キャッシュは C:\Users\john\AppData\Local\Microsoft\Office\16.0\WEF にある場合があります。

> [!IMPORTANT]
> Windows 上の Outlook 2013 では、16.0 の代わりに 15.0 を使用します。場所は次のようになります。
>
> ```text
> %LocalAppData%\Microsoft\Office\15.0\WEF
> ```

アドインがどのアイテムに対してもアクティブ化されない場合、マニフェストが Exchange Server 上に適切にインストールされなかったか、あるいは、Outlook が起動時に正しくマニフェストを読み取れなかった可能性があります。Exchange 管理センターを使用して、アドインがメールボックスにインストールされ、有効化されていることを確認し、必要に応じて Exchange Server を再起動します。

図 1 は、Outlook に有効なバージョンのマニフェストがあるかどうかを確認するステップの概要を示しています。

**図 1.Outlook がマニフェストを適切にキャッシュしたかどうかを確認するステップのフローチャート**

![マニフェストをチェックするフローチャート。](../images/troubleshoot-manifest-flow.png)

以下の手順では、その詳細を説明します。

1. Outlook を開いている間にマニフェストを変更し、アドインの開発に Visual Studio 2012 や Visual Studio の新しいバージョンを使用していない場合は、Exchange 管理センターを使用して、そのアドインをアンインストールし、再インストールする必要があります。

1. Outlook を再起動し、Outlook でアドインがアクティブになっているかどうかをテストします。

1. アドインがアクティブ化されない場合は、アドインのマニフェストの適切なキャッシュ コピーが Outlook にあるかどうかを確認します。 次のパスの下を見てください。

    ```text
    %LocalAppData%\Microsoft\Office\16.0\WEF
    ```

    マニフェストは、次のサブフォルダーにあります。

    ```text
    \<insert your guid>\<insert base 64 hash>\Manifests\<ManifestID>_<ManifestVersion>
    ```

    > [!NOTE]
    > ユーザー John のメールボックスにインストールされたマニフェストへのパスの例を次に示します。
    >
    > ```text
    > C:\Users\john\appdata\Local\Microsoft\Office\16.0\WEF\{8D8445A4-80E4-4D6B-B7AC-D4E6AF594E73}\GoRshCWa7vW8+jhKmyiDhA==\Manifests\b3d7d9d5-6f57-437d-9830-94e2aaccef16_1.2
    > ```

    テストしているアドインのマニフェストが、キャッシュされたマニフェストに含まれているかどうかを確認します。

1. マニフェストがキャッシュにある場合は、このセクションの残りの部分をスキップして、このセクションの後で説明している、他に考えられる理由を検討します。

1. マニフェストがキャッシュにない場合は、Outlook が Exchange Server から実際にマニフェストを読み取ったかどうかを確認します。これを行うには、Windows イベント ビューアーを使用します。

    1. **[Windows ログ]** で **[アプリケーション]** を選択します。

    1. イベント ID が 63 に等しい比較的最近のイベントを探します。これは、Outlook が Exchange Server からマニフェストをダウンロードしたことを表します。

    1. Outlook がマニフェストを正常に読み取った場合、ログに記録されたイベントには次の説明が必要です。

        ```text
        The Exchange web service request GetAppManifests succeeded.
        ```

        このセクションの残りの部分をスキップして、このセクションの後で説明している、他に考えられる理由を検討します。

1. 成功したイベントが表示されない場合は、Outlook を閉じて、次のパス内のすべてのマニフェストを削除します。

    ```text
    %LocalAppData%\Microsoft\Office\16.0\WEF\<insert your guid>\<insert base 64 hash>\Manifests\
    ```

    Outlook を起動し、Outlook でアドインがアクティブになっているかどうかをテストします。

1. アドインがアクティブ化されない場合は、手順 3 に戻り、Outlook がマニフェストを適切に読み取ったかどうかを再度確認します。

## <a name="is-the-add-in-manifest-valid"></a>アドイン マニフェストは有効か?

「[マニフェストの問題を検証し、トラブルシューティングを行う](../testing/troubleshoot-manifest.md)」を参照して、アドインのマニフェストの問題をデバッグしてください。

## <a name="are-you-using-the-appropriate-activation-rules"></a>適切なアクティブ化ルールを使用しているか?

Office アドイン マニフェスト スキーマ バージョン 1.1 以降では、ユーザーが新規作成フォームを使用しているときにアクティブ化されるアドイン (新規作成アドイン) や閲覧フォームを使用しているときにアクティブ化されるアドイン (閲覧アドイン) を作成できます。アドインをアクティブ化するフォームの種類に適した正しいアクティブ化ルールを指定してください。たとえば、新規作成アドインをアクティブ化する場合は、[FormType](/javascript/api/manifest/rule#itemis-rule) 属性が **Edit** または **ReadOrEdit** に設定された **ItemIs** ルールのみを使用する必要があり、[ItemHasKnownEntity](/javascript/api/manifest/rule#itemhasknownentity-rule) ルールや [ItemHasRegularExpressionMatch](/javascript/api/manifest/rule#itemhasregularexpressionmatch-rule) ルールなど他の型のルールを新規作成アドイン用に使用することはできません。詳細については、「[Outlook アドインのアクティブ化ルール](activation-rules.md)」を参照してください。

## <a name="if-you-use-a-regular-expression-is-it-properly-specified"></a>正規表現を使用している場合、正しく指定されていますか。

アクティブ化ルール内の正規表現は閲覧アドインの XML マニフェスト ファイルの一部であるため、正規表現で特定の文字を使用する場合は、XML プロセッサがサポートする対応するエスケープ シーケンスに従う必要があります。表 1 にこのような特殊文字を示します。

**表 1.正規表現のエスケープ シーケンス**

|**文字**|**説明**|**使用するエスケープ シーケンス**|
|:-----|:-----|:-----|
|`"`|二重引用符|&amp;quot;|
|`&`|アンパサンド|&amp;amp;|
|`'`|アポストロフィ|&amp;apos;|
|`<`|より小さい|&amp;lt;|
|`>`|より大きい|&amp;gt;|

## <a name="if-you-use-a-regular-expression-is-the-read-add-in-activating-in-outlook-on-the-web-or-mobile-devices-but-not-in-any-of-the-outlook-rich-clients"></a>正規表現を使用する場合、閲覧アドインは Outlook on the web またはモバイル デバイスではアクティブ化されるものの、どの Outlook リッチ クライアントでもアクティブ化されないか?

Outlook リッチ クライアントでは、Outlook on the web とモバイル デバイスで使用されている正規表現エンジンとでは、異なる正規表現エンジンを使用します。Outlook リッチ クライアントでは、Visual Studio の標準テンプレート ライブラリの一部として提供されている C++ 正規表現エンジンを使用します。このエンジンは ECMAScript 5 標準に準拠しています。Outlook on the web およびモバイル デバイスでは、JavaScript の一部である正規表現評価を使用します。これはブラウザーによって提供されるものであり、ECMAScript 5 のスーパーセットをサポートしています。

ほとんどの場合、これらの Outlook クライアントは、アクティブ化ルールで同じ正規表現に対して同じ一致を見つけますが、例外があります。 たとえば、正規表現に定義済みの文字クラスに基づくカスタム文字クラスが含まれている場合、Outlook リッチ クライアントはOutlook on the webやモバイル デバイスとは異なる結果を返す可能性があります。 たとえば、文字クラス内に短縮形の文字クラス `[\d\w]` が含まれていると、結果にばらつきが生じます。 この場合、異なるアプリケーションで異なる結果を回避するには、代わりに使用 `(\d|\w)` します。

正規表現を十分にテストしてください。異なる結果が返された場合は、両方のエンジンでの互換性のために正規表現を書き換えます。Outlook リッチ クライアントの評価結果を確認するには、一致させるテキストのサンプルに対して正規表現を適用させる小さな C++ プログラムを作成します。Visual Studio 上で動作する C++ テスト プログラムは、標準テンプレート ライブラリを使用して、同じ正規表現を実行しているときに Outlook リッチ クライアントの動作をシミュレートします。Outlook on the web またはモバイル デバイスでの評価結果を確認するには、お好きな JavaScript 正規表現テスターを使用してください。

## <a name="if-you-use-an-itemis-itemhasattachment-or-itemhasregularexpressionmatch-rule-have-you-verified-the-related-item-property"></a>ItemIs ルール、ItemHasAttachment ルール、または ItemHasRegularExpressionMatch ルールを使用する場合、関連するアイテム プロパティを確認しましたか。

**ItemHasRegularExpressionMatch** アクティブ化ルールを使用する場合は、**PropertyName** 属性の値が、選択されているアイテムの予期する値かどうかを確認します。 対応するプロパティをデバッグするためのヒントを次に示します。

- 選択されているアイテムがメッセージであり、**PropertyName** 属性に **BodyAsHTML** を指定する場合は、メッセージを開いて **[ソースの表示]** を選択し、そのアイテムの HTML 表現でのメッセージ本文を確認します。

- 選択されているアイテムが予定の場合、またはアクティブ化ルールで **PropertyName** に **BodyAsPlaintext** が指定される場合は、Windows での Outlook で Outlook オブジェクト モデルと Visual Basic エディターを使用できます。

    1. マクロが有効で、**[開発]** タブが Outlook のリボンに表示されていることを確認します。

    1. Visual Basic エディターで、**[表示]**、**[イミディエイト ウィンドウ]** を選択します。

    1. シナリオに応じて各種のプロパティを表示するには、次のように入力します。

        - Outlook エクスプローラーで選択されているメッセージ アイテムまたは予定アイテムの HTML 形式の本文。

        ```vb
        ?ActiveExplorer.Selection.Item(1).HTMLBody
        ```
        - Outlook エクスプローラーで選択されているメッセージ アイテムまたは予定アイテムのプレーン テキスト形式の本文。

        ```vb
        ?ActiveExplorer.Selection.Item(1).Body
        ```
        - 現在の Outlook インスペクターで開かれているメッセージ アイテムまたは予定アイテムの HTML 形式の本文。

        ```vb
        ?ActiveInspector.CurrentItem.HTMLBody
        ```
        - 現在の Outlook インスペクターで開かれているメッセージ アイテムまたは予定アイテムのプレーン テキスト形式の本文。

        ```vb
        ?ActiveInspector.CurrentItem.Body
        ```

**ItemHasRegularExpressionMatch** アクティブ化ルールで **Subject** または **SenderSMTPAddress** が指定される場合、あるいは **ItemIs** ルールまたは **ItemHasAttachment** ルールを使用していて、MAPI の使用に精通しているか使用する必要がある場合は、[MFCMAPI](https://github.com/stephenegriffin/mfcmapi) を使用して、ルールで使用される表 2 の値を確認できます。

**表 2アクティブ化ルールと対応する MAPI プロパティ**

|ルールの種類|確認する MAPI プロパティ|
|:-----|:-----|
|**ItemHasRegularExpressionMatch** ルールと **Subject**|[PidTagSubject](/office/client-developer/outlook/mapi/pidtagsubject-canonical-property)|
|**ItemHasRegularExpressionMatch** ルールと **SenderSMTPAddress**|
  [PidTagSenderSmtpAddress](/office/client-developer/outlook/mapi/pidtagsendersmtpaddress-canonical-property) と [PidTagSentRepresentingSmtpAddress](/office/client-developer/outlook/mapi/pidtagsentrepresentingsmtpaddress-canonical-property)|
|**ItemIs**|[PidTagMessageClass](/office/client-developer/outlook/mapi/pidtagmessageclass-canonical-property)|
|**ItemHasAttachment**|[PidTagHasAttachments](/office/client-developer/outlook/mapi/pidtaghasattachments-canonical-property)|

プロパティ値を確認した後、正規表現評価ツールを使用して、正規表現がその値の中で一致を見つけるかどうかをテストできます。

## <a name="does-outlook-apply-all-the-regular-expressions-to-the-portion-of-the-item-body-as-you-expect"></a>Outlook では、期待どおりにアイテム本体の部分にすべての正規表現が適用されますか?

このセクションは、正規表現を使用するすべてのアクティブ化規則 (特にアイテム本文に適用されるもの) に適用されます。これはサイズが大きく、一致の評価に時間がかかる場合があります。 アクティブ化ルールが依存する item プロパティに期待した値が含まれている場合でも、Outlook ではアイテム プロパティの値全体のすべての正規表現を評価できない可能性があることに注意してください。 適切なパフォーマンスを提供し、読み取りアドインによる過度のリソース使用量を制御するために、Outlook では、実行時にアクティブ化規則で正規表現を処理する際に次の制限が適用されます。

- 評価されるアイテム本文のサイズ - Outlook が正規表現を評価するアイテム本文の部分に制限があります。 これらの制限は、Outlook クライアント、フォーム ファクター、アイテム本文の形式によって異なります。 詳細については、「[Outlook アドインのアクティブ化と JavaScript API の制限](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)」の表 2 を参照してください。

- 正規表現の一致の数 - Outlook リッチ クライアント、Outlook on the web、モバイル デバイスは、それぞれ正規表現の一致を 50 件まで返します。これらの一致は一意であり、重複の一致はこの制限にカウントされません。返される一致の順序を想定しないでください。Outlook リッチ クライアントでの順序は Outlook on the web およびモバイル デバイスでの順序と同じとは限りません。アクティブ化ルールに正規表現の一致が多数存在することが予想されるにもかかわらず、一致が見つからない場合は、この制限を超えている可能性があります。

- 正規表現の一致の長さ -- Outlook アプリケーションが返す正規表現一致の長さに制限があります。 Outlook には、制限を超えて一致するものは含まれていないので、警告メッセージは表示されません。 他の regex 評価ツールまたはスタンドアロンの C++ テスト プログラムで正規表現を実行して、このような制限を超える一致があるかどうかを確認できます。 表 3 にこの制限の要約を示します。 詳細については、「[Outlook アドインのアクティブ化と JavaScript API の制限](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)」の表 3 を参照してください。

    **表 3正規表現の一致の長さ制限**

    |正規表現の長さ制限|Outlook リッチ クライアント|Outlook on the web またはモバイル デバイス|
    |:-----|:-----|:-----|
    |アイテムの本文がテキスト形式の場合|1.5 KB|3 KB|
    |アイテムの本文が HTML の場合|3 KB|3 KB|

- Outlook リッチ クライアント用閲覧アドインのすべての正規表現の評価にかかった時間 : 既定では、Outlook はアクティブ化ルール内のすべての正規表現の評価を閲覧アドインごとに 1 秒以内で完了する必要があります。完了しなかった場合、Outlook は最大 3 回まで再試行し、それでも評価を完了できないとアドインを無効化します。Outlook は、アドインが無効になったというメッセージを通知バーに表示します。正規表現に使用可能な時間の長さは、グループ ポリシーまたはレジストリ キーの設定で変更できます。 

   > [!NOTE]
   > Outlook リッチ クライアントが、閲覧アドインを無効にした場合、閲覧アドインは、Outlook リッチ クライアント、Outlook on the web、モバイル デバイスの同じメールボックスで使用できなくなります。

## <a name="see-also"></a>関連項目

- [テスト用に Outlook アドインを展開してインストールする](testing-and-tips.md)
- [Outlook アドインのアクティブ化ルール](activation-rules.md)
- [正規表現アクティブ化ルールを使用して Outlook アドインを表示する](use-regular-expressions-to-show-an-outlook-add-in.md)
- [Outlook アドインのアクティブ化と JavaScript API の制限](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [マニフェストの問題を検証し、トラブルシューティングする](../testing/troubleshoot-manifest.md)
