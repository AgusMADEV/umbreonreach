<?php 

    // Process Logout
    if (isset($_GET['action']) && $_GET['action'] === 'logout'){
        session_destroy();
        header("Location: index.php");
        exit;
    }

?>