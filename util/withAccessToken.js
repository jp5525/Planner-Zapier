const requestWithAccessToken = (req, z, bundle) => {
    
    const {appId, clientSecret, tenant} = bundle.authData;
    
    return z.request(`https://login.microsoftonline.com/`+ tenant +`/oauth2/v2.0/token`, {
        method: 'POST',
        body: {
            client_id: appId,
            scope: "https://graph.microsoft.com/.default",
            client_secret: clientSecret,
            grant_type: 'client_credentials',
        },
        headers: {
            'content-type': 'application/x-www-form-urlencoded'
        }
    })
    .then(response=>{
        //z.console.log(response.json)
        const header = {...(req.headers || {}), ...{Authorization :`Bearer ${response.json.access_token}`}};
       // z.console.log("--------------------------------")
        //z.console.log({...req, ...{headers:header}})
        
        return z.request({...req, ...{headers:header}}).catch(err=>console.log(JSON.stringify(req,null, 2)))
    })
      
  };

      
module.exports = { requestWithAccessToken }
  