function chkConfirmDelete(){
	if(confirm("確定要刪除項目嗎？")){return true;}
	return false ;
}

function getAttachment(ProjectID,DataID,Mode,AttachmentType){
		var theURL = "attachment.aspx?ProjectID=" + ProjectID + "&RecordID=" +  DataID + "&Mode=" + Mode  + "&AttachmentType=" + AttachmentType  + "&AttachmentFiles=" + document.getElementById("txtFilename").value ;
		var CWIN=window.open(theURL,"PM_attachment","toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=auto,resizable=1" ,true)
	}
	
function DeleteReply(topicID,replyID)	
{
	if(confirm("確定要刪除回覆嗎？")){
	document.location.href="ForumTopicDetail.aspx?TopicID=" + topicID + "&ProjectID=1&index=-1&projectIndex=6&DelReplyID=" + replyID ;
	}
}