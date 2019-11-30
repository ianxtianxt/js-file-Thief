<?php
if($_SERVER['REQUEST_METHOD'] == 'POST'){
	$fileName = uniqid(rand()) . '_' . iconv('utf-8', 'gbk', $_SERVER['HTTP_FILENAME']);
	print_r(file_put_contents("uploads/{$fileName}", $HTTP_RAW_POST_DATA));
}
?>