function onSignIn(googleUser){
	
	
		var profile=googleUser.getBasicProfile();
		$(".g-signin2").css("display","none");
		$(".data").css("display","block");
		$("#pic").attr('src',profile.getImageUrl());
		$("#email").text(profile.getEmail());
		
	
		
		location.replace("http://localhost:8080/My%20Programs/Web%20Programming%20Project/final/homePage.html");


		
}


