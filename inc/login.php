<?php

    // Process login
    if (isset($_POST['action']) && $_POST['action'] === 'login'){
        $username = trim($_POST['username'] ?? '');
        $password = trim($_POST['password'] ?? '');

        $stmt = $db->prepare("
            SELECT * FROM users
            WHERE username = :u AND password = :p
            LIMIT 1
        ");
        $stmt->execute([':u' => $username, 'p' => $password]);
        $user = $stmt->fetch(PDO::FETCH_ASSOC);

        if ($user){
            $_SESSION['user_id'] = $user['id'];
            $_SESSION['username'] = $user['username'];
            header("Location: index.php");
            exit;
        }else{
            $loginError = "Credenciales inválidas.";
        }
    }

?>