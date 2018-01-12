
# <a name="use-the-oauth-authorization-framework-in-an-office-add-in"></a>Office アドインで OAuth 認証フレームワークを使用する

OAuth は、Office 365、Facebook、Google、SalesForce、LinkedIn などのオンライン サービス プロバイダーがユーザー認証を実行するのに使用する認証のオープン標準です。OAuth 認証フレームワークは、Azure と Office 365 で使用される既定の認証プロトコルです。OAuth 認証フレームワークは、エンタープライズ (企業) とコンシューマー シナリオの両方で使用されます。

オンライン サービス プロバイダーは、REST 経由で公開されているパブリックの API を提供することがあります。開発者は、オンライン サービス プロバイダーにデータを読み書きするために、Office アドインでこれらのパブリック API を使用できます。アドインでオンライン サービス プロバイダーからのデータを統合することにより、その価値が高められ、ユーザーによる採用状況が向上します。アドインでこれらの API を使用する場合、ユーザーは OAuth 認証フレームワークを使用して認証することが必要です。

このトピックでは、アドインで認証フローを実装して、ユーザー認証を実行する方法について説明します。このトピックに含まれているコード セグメントは、[Office-Add-in-NodeJS-ServerAuth](https://github.com/OfficeDev/Office-Add-in-NodeJS-ServerAuth) のコード サンプルから採用されています。

 **メモ**  セキュリティ上の理由から、ブラウザーは IFrame のサインイン ページを表示できません。お客様が使用している Office のバージョンによって、特に Web ベースのバージョンによっては、アドインは IFrame で表示されます。これは、認証フローを管理する方法におけるいくつかの考慮事項を提起します。 

次のダイアグラムは、必要なコンポーネントと、アドインで認証を実装するときに発生するイベントのフローを示します。

![Office アドインでの OAuth 認証の実行](../../images/OAuthInOfficeAddin.png)

ダイアグラムは、次の必要なコンポーネントを使用する方法を示しています。


- Office は作業ウィンドウ アドインをユーザーのコンピューター上で実行します。アドインは、認証フローを開始するためのポップアップ ウィンドウを開きます。使用するプラットホームによりますが、アドインは IFRAME で実行される可能性があるため、アドインが認証フローを直接開始することはできません。セキュリティ上の理由から、OAuth サインイン ページは IFRAME に表示できません。 
    
- Web サーバーはアドインのコードをホストします。このコード サンプルでは、ユーザーのアクセス トークンを格納するために Web サーバーで実行されているデータベース サーバーを使用します。アクセス トークンの保持が必要です。そうすれば、ポップアップ ウィンドウを使用して認証が完了した後、メインのアドインのページが同じトークンを使用してオンライン サービスからのデータにアクセスできます。アドインまたはポップアップから渡される情報に依存できないため、サーバー側のオプションを使用してトークンを保存することが必要です。
    
- OAuth 2.0 プロバイダーはユーザー認証を行います。
    

    
 **重要**  アクセス トークンは作業ウィンドウに返すことはできませんが、サーバー上で使用することができます。このコード サンプルでは、アクセス トークンはデータベースに 2 分間保存されます。2 分後、トークンは、データベースから削除され、ユーザーは再認証するよう求められます。独自の実装でこの期間を変更する前に、2 分よりも長い期間データベースにアクセス トークンを格納する場合のセキュリティ上のリスクを考慮してください。


## <a name="step-1---start-socket-and-open-a-pop-up-window"></a>手順 1 - ソケットを開始してポップアップ ウィンドウを開く

このコード サンプルを実行すると、Office で作業ウィンドウ アドインが表示されます。ユーザーがログインするのに OAuth プロバイダーを選択すると、アドインはまずソケットを作成します。このサンプルでは、アドインで優れたユーザー エクスペリエンスを提供するためにソケットを使用します。アドインは、ユーザーに認証の成否を伝達するためにソケットを使用します。ソケットを使用すれば、アドインのメイン ページは認証状態で簡単に更新され、ユーザーの操作またはポーリングを必要としません。routes/connect.js から取られた次のコード セグメントでは、ソケットを開始する方法を示します。ソケットには、アドインのセッション ID である **decodedNodeCookie** を使用して名前を付けます。このコード サンプルは、[socket.io](http://socket.io/) を使用してソケットを作成します。


```js
io.on('connection', function (socket) {
  console.log('Socket connection established');
  var jsonCookie =
    cookie.parse(socket
      .handshake
      .headers
      .cookie);
  var decodedNodeCookie =
    cookieParser
      .signedCookie(jsonCookie.nodecookie, '<Insert a random string>');
  console.log('Decoded cookie: ' + decodedNodeCookie);
  // The session ID becomes the room name for this session.
  socket.join(decodedNodeCookie);
  io.to(decodedNodeCookie).emit('init', 'Private socket session established');
});

```

次に、アドインはソケットに接続します。次のコードが /public/javascripts/client.js にあります。




```js
var socket = io.connect('https://localhost:3001', { secure: true });
```

次に、アドインは、**window.open** を使用してユーザーのコンピューター上でポップアップ ウィンドウを開きます。**window.open** を実行する場合、リダイレクト URI とアドインのセッション ID が URL で渡されるようにします。アドインのセッション ID は、アドインの UI に認証状態情報を送信するときに、使用するソケットを指定するために使用されます。次のコード セグメントが、views/index.jade にあります。




```js
onclick="window.open('/connect/azure/#{sessionID}', 'AuthPopup', 'width=500,height=500,centerscreen=1,menubar=0,toolbar=0,location=0,personalbar=0,status=0,titlebar=0,dialog=1')")
```


## <a name="steps-2-amp-3---start-the-authentication-flow-and-show-the-sign-in-page"></a>手順 2 と 3 - 認証フローを開始して、サインイン ページを表示する

アドインは認証フローを開始する必要があります。次のコード セグメントでは、Passport OAuth ライブラリを使用します。認証フローを開始するときは、OAuth プロバイダーの認証 URL と、アドインのセッション ID を渡すことを確認します。アドインのセッション ID は、State パラメーターで渡す必要があります。ポップアップ ウィンドウに、ユーザーがサインインするための OAuth プロバイダーのサインイン ページが表示されます。


```js
router.get('/azure/:sessionID', function(req, res, next) { 
   passport.authenticate( 
     'azure',  
     { state: req.params.sessionID }, 

```


## <a name="steps-4-5-amp-6---user-signs-in-and-web-server-receives-tokens"></a>手順 4、5、6 - ユーザーがサインインして、Web サーバーでトークンを受信する

 正常にサインインした後で、アクセス トークン、リフレッシュ トークン、State パラメーターがアドインに返されます。State パラメーターには、セッション ID が含まれます。これは、手順 7 でソケットへ認証状態情報を送信するために使用されます。app.js から取られた次のコード セグメントは、データベースにアクセス トークンを格納します。


```js
  dbHelperInstance.insertDoc(userData, null, 
         function (err, body) { 
           if (!err) { 
             console.log("Inserted session entry [" + userData.sessid + "] id: " + body.id); 
           } 
           done(err, userData); 
         }); 

```


## <a name="step-7---show-authentication-information-in-the-add-ins-ui"></a>手順 7 - アドインの UI で認証情報を表示する

connect.js から取られた次のコード セグメントは、認証状態情報を使用してアドインの UI を更新します。手順 1 で作成されたソケットを使用して、アドインの UI が更新されます。


```js
  
       io.to(user.sessid).emit('auth_success', providers); 
       next(); 

```


## <a name="additional-resources"></a>その他のリソース
<a name="bk_addresources"> </a>


- [Office アドインのサーバー認証サンプル Node.js 用](https://github.com/OfficeDev/Office-Add-in-Nodejs-ServerAuth/blob/master/README.md)
    
