---
ms.date: 1/29/2019
description: Excel でカスタム関数を使用してユーザーを認証します。
title: ユーザー定義関数での認証
ms.openlocfilehash: 0e42dbc93cb545660a8dbaae5bdb48724f3b7376
ms.sourcegitcommit: 33dcf099c6b3d249811580d67ee9b790c0fdccfb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/05/2019
ms.locfileid: "29745418"
---
# <a name="authentication"></a>認証

保護されたリソースを場合によっては、ユーザー定義関数にアクセスするためにユーザーを認証する必要があります。 カスタム関数は、特定の認証方法を必要としない、作業ウィンドウと、アドインの場合は、他の UI 要素から別の実行時にユーザー定義関数を実行するに注意する必要があります。 使用して 2 つのランタイムの間で前後にデータを渡す必要があります、このため、`AsyncStorage`オブジェクトとダイアログ ボックス API です。
  
## <a name="asyncstorage-object"></a>AsyncStorage オブジェクト

ユーザー定義関数の実行時の有効期限がない、`localStorage`される可能性があります通常データを格納するグローバル ウィンドウで、使用可能なオブジェクトです。 代わりに、設定およびデータを取得する[OfficeRuntime.AsyncStorage](https://docs.microsoft.com/javascript/api/office-runtime/officeruntime.asyncstorage)を使用して、独自の機能と作業ウィンドウ間でデータを共有する必要があります。 

使用するメリットがあるさらに、 `AsyncStorage`。その他のアドインを使用して、データにアクセスできないようにセキュリティで保護されたサンド ボックス環境を使用します。  

### <a name="suggested-usage"></a>推奨される使用方法

作業ウィンドウまたはカスタム関数のいずれかを認証する場合は、アクセス トークンがすでに取得したかどうかを参照してくださいに AsyncStorage を確認してください。 それ以外の場合は、ダイアログ ボックス API を使用して、ユーザーを認証し、アクセス トークンを取得し、AsyncStorage で後で使用できるトークンを格納します。

## <a name="dialog-api"></a>ダイアログ API

トークンが存在しない場合は、サインインするユーザーを確認するダイアログ ボックス API を使用してください。 作成されたアクセス トークンを格納できるユーザーが各自の資格情報を入力した後`AsyncStorage`。

> [!NOTE]
> ユーザー定義関数の実行時では、作業ウィンドウで使用される実行時にダイアログ オブジェクトとは少し異なりますダイアログ オブジェクトを使用します。 いる両方と呼ばれる「ダイアログ API」、これらの使用が`Officeruntime.Dialog`ユーザー定義関数の実行時にユーザー認証します。

使用する方法については、 `OfficeRuntime.Dialog`、[実行時のカスタム関数](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-runtime?view=office-js#displaying-a-dialog-box)を参照してください。

全体として全体の認証プロセスを予見するには場合があります] 作業ウィンドウと、アドインの UI 要素と考えるとよいとカスタムを通じて相互に通信できる、個別のエンティティとして、アドインの一部の機能`AsyncStorage`。

次の図では、この基本的な手順について説明します。 点線に個別の操作を実行すると、ユーザー定義関数、アドインの作業ウィンドウは、アドインを全体としての両方の部分を示すことに注意してください。

1. Excel ブック内のセルからユーザー定義関数の呼び出しを発行するとします。
2. ユーザー定義関数を使用して`Officeruntime.Dialog`web サイトにユーザーの資格情報を渡すことです。
3. この web サイトは、アクセス トークンをユーザー定義関数に戻ります。
4. ユーザー定義関数が、このアクセス トークンを設定、 `AsyncStorage`。
5. アドインの作業ウィンドウからのトークンにアクセスする`AsyncStorage`。

![ユーザー定義関数、OfficeRuntime、および共同作業の作業ウィンドウのダイアグラム]。(../images/Authdiagram.png "認証のダイアグラム")。

## <a name="general-guidance"></a>一般的なガイダンス

Office アドインでは、web ベースおよび web のすべての認証テクニックを使用することができます。 特定のパターンまたはカスタム関数を使用して、独自の認証を実装するメソッドはありません。 できます各種の認証パターンについてのマニュアルを参照する[外部サービスを使用して付与するには、この資料](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins?view=office-js)で始まります。  

次の場所を使用してカスタム機能を開発するときにデータを格納するを回避するには。  

- `localStorage`: ユーザー定義関数では、グローバルへのアクセスを必要はありません`window`オブジェクト、に格納されているデータへのアクセスはそのためありません`localStorage`。
- `Office.context.document.settings`: この場所は安全ではありませんし、アドインを使用するすべてのユーザーが情報を抽出することができます。

## <a name="see-also"></a>関連項目

* [カスタム関数のメタデータ](custom-functions-json.md)
* [Excel カスタム関数のランタイム](custom-functions-runtime.md)
* [カスタム関数のベスト プラクティス](custom-functions-best-practices.md)
* [チュートリアル: Excel でカスタム関数を作成します。](excel-tutorial-custom-functions.md)
