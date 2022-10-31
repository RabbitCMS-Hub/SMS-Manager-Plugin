function SMSTestResult(vText){
	$('#SMSTestSonucu').html(`<pre class="alert alert-info mt-2" style="white-space:break-spaces;color:#fff;">${vText}</pre>`);
	$('#SendSMSPreview').prop('disabled', false);	
};