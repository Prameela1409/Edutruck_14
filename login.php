<?php
// Replace these values with your actual database credentials
$servername = "localhost";
$username = "root";
$password = "";
$dbname = "signin";

// Create a connection to the database
$conn = new mysqli($servername, $username, $password, $dbname);

// Check the connection
if ($conn->connect_error) {
    die("Connection failed: " . $conn->connect_error);
}

// Get data from the form
$email = $_POST['email'];
$password = $_POST['password'];

// Check the credentials in the database
$sql = "SELECT * FROM signin WHERE email = '$email'";
$result = $conn->query($sql);

if ($result->num_rows > 0) {
    $row = $result->fetch_assoc();
    if (password_verify($password, $row['password'])) {
        // Successful login
        session_start();
        $_SESSION['email'] = $email;
        header("Location: edutrack.html");
        exit();
    } else {
        echo "Invalid password";
    }
} else {
    echo "User not found";
}

// Close the database connection
$conn->close();
?>
