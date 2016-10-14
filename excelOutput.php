<?php

header("Content-type: application/vnd.ms-excel"); //文件內容為excel格式
header("Content-Disposition: attachment; filename=osTicket".date("Ymd").".xls;"); //將PHP轉成下載的檔案指定名稱與副檔名.xls
 
echo '<HTML xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">'."\n";
echo '<head><meta http-equiv="content-type" content="application/vnd.ms-excel; charset=UTF-8"></head>'."\n";
echo '<body>';
//header
echo '<table>
	<tr>
		<td>ticket_id</td>
		<td>ticketID</td>
		<td>dept_id</td>
		<td>sla_id</td>
		<td>priority_id</td>
		<td>topic_id</td>
		<td>staff_id</td>
		<td>team_id</td>
		<td>email</td>
		<td>name</td>
		<td>subject</td>
		<td>phone</td>
		<td>phone_ext</td>
		<td>OS</td>
		<td>product</td>
		<td>serial</td>
		<td>subTopic</td>
		<td>ip_address</td>
		<td>status</td>
		<td>source</td>
		<td>isoverdue</td>
		<td>isanswered</td>
		<td>duedate</td>
		<td>reopened</td>
		<td>closed</td>
		<td>lastmessage</td>
		<td>lastresponse</td>
		<td>created</td>
		<td>updated</td>
	</tr>';
//Link Sql
$link = mysql_connect('localhost', 'CustomerService', 'IEC020201') or die('Could not connect: ' . mysql_error());
mysql_query("SET NAMES 'utf8'");
mysql_select_db('customerservice');
$sql = "SELECT * FROM `ost_ticket`;";
$res = mysql_query($sql);
while($row = mysql_fetch_array($res)){
	$ticket_id = $row['ticket_id'];
	$ticketID = $row['ticketID'];
	$dept_id = $row['dept_id'];
	$sla_id = $row['sla_id'];
	$priority_id = $row['priority_id'];
	$topic_id = $row['topic_id'];
	$staff_id = $row['staff_id'];
	$team_id = $row['team_id'];
	$email = $row['email'];
	$name = $row['name'];
	$subject = $row['subject'];
	$phone = $row['phone'];
	$phone_ext = $row['phone_ext'];
	$OS = $row['OS'];
	$product = $row['product'];
	$serial = $row['serial'];
	$subTopic = $row['subTopic'];
	$ip_address = $row['ip_address'];
	$status = $row['status'];
	$source = $row['source'];
	$isoverdue = $row['isoverdue'];
	$isanswered = $row['isanswered'];
	$duedate = $row['duedate'];
	$reopened = $row['reopened'];
	$closed = $row['closed'];
	$lastmessage = $row['lastmessage'];
	$lastresponse = $row['lastresponse'];
	$created = $row['created'];
	$updated = $row['updated'];
	echo '<table>
		<tr>
			<td>'.$ticket_id.'</td>
			<td>'.$ticketID.'</td>
			<td>'.$dept_id.'</td>
			<td>'.$sla_id.'</td>
			<td>'.$priority_id.'</td>
			<td>'.$topic_id.'</td>
			<td>'.$staff_id.'</td>
			<td>'.$team_id.'</td>
			<td>'.$email.'</td>
			<td>'.$name.'</td>
			<td>'.$subject.'</td>
			<td>'.$phone.'</td>
			<td>'.$phone_ext.'</td>
			<td>'.$OS.'</td>
			<td>'.$product.'</td>
			<td>'.$serial.'</td>
			<td>'.$subTopic.'</td>
			<td>'.$ip_address.'</td>
			<td>'.$status.'</td>
			<td>'.$source.'</td>
			<td>'.$isoverdue.'</td>
			<td>'.$isanswered.'</td>
			<td>'.$duedate.'</td>
			<td>'.$reopened.'</td>
			<td>'.$closed.'</td>
			<td>'.$lastmessage.'</td>
			<td>'.$lastresponse.'</td>
			<td>'.$created.'</td>
			<td>'.$updated.'</td>
		</tr>';
}
	echo '
	</table>
	';
 
echo '</body></html>';
?>