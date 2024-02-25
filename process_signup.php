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
$password = password_hash($_POST['password'], PASSWORD_DEFAULT); // Hash the password for security

// Insert data into the database
$sql = "INSERT INTO signin (email, password) VALUES ('$email', '$password')";

if ($conn->query($sql) === TRUE) {
    // After successful signup
    echo "Signup successful!";
    header("Location: loginsf.html");
    exit();
} else {
    echo "Error: " . $sql . "<br>" . $conn->error;
}
// After successful signup
echo "Signup successful!";
header("Location: loginsf.html");
exit();

// Close the database connection
$conn->close();
?>

