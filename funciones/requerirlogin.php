<?php

    function requiereLogin(){
        if (!isset($_SESSION['user_id'])){
            header("Location: index.php");
        }
    }

?>