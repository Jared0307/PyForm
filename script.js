document.addEventListener('DOMContentLoaded', function() {
    const form = document.querySelector('form');
    const errorMessages = document.createElement('div');
    errorMessages.className = 'error';
    form.prepend(errorMessages);

    form.addEventListener('submit', function(event) {
        let hasError = false;
        errorMessages.textContent = '';

        form.querySelectorAll('input[type="text"]').forEach(input => {
            if (input.value.trim() === '') {
                hasError = true;
                errorMessages.textContent += `El campo ${input.previousElementSibling.textContent.trim()} es obligatorio.\n`;
            }
        });

        if (hasError) {
            event.preventDefault();
            return false;
        }
    });
});
