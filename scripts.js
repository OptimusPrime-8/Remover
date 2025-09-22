// Get the button element
const changeColorButton = document.getElementById('changeColorButton');

// Add an event listener for the button click
changeColorButton.addEventListener('click', function() {
    // Generate a random color
    const randomColor = `hsl(${Math.random() * 360}, 100%, 75%)`;
    // Change the background color of the page
    document.body.style.backgroundColor = randomColor;
});
