const router =
require(
"express"
)
.Router();



router.post(

"/login",

(req,res)=>{

const {

username,

password

}
=
req.body;



if(

username==="1234"

&&

password==="1234"

){

/* CREATE SESSION */

req.session.loggedIn =
true;



res.json({

success:true,

redirect:
"/dashboard"

});

}

else{

res.json({

success:false

});

}

}



);



module.exports =
router;