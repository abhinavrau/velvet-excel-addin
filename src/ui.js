



export function showStatus(message, isError) {
    $('.status').empty();
    
    var element = $('<div/>', {
        class: `status-card ms-depth-4 ${isError ? 'error-msg' : 'success-msg'}`
    }).append($('<p/>', {
        class: 'ms-fontSize-24 ms-fontWeight-bold',
        text: isError ? 'An error occurred' : 'Success',
        class: 'ms-fontSize-16 ms-fontWeight-regular',
        text: message
    }));
        
    $('.status').append(element);
}