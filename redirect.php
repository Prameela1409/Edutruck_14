<?php
// Check if the 'page' parameter is set in the URL
if (isset($_GET['page'])) {
    // Get the value of the 'page' parameter
    $page = $_GET['page'];
    
    // Redirect to the appropriate page based on the value of the 'page' parameter
    if ($page === 'neededcourses') {
        header("Location: neededcources.html");
        exit();
    } else {
        // Handle other cases or invalid values
        echo "Invalid page parameter.";
    }
} else {
    // If the 'page' parameter is not set, display an error message
    echo "Page parameter is missing.";
}
?>
