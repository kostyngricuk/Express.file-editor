const uploadForm = document.getElementById('upload-form')
uploadForm.addEventListener('submit', uploadFormSubmit)

function uploadFormSubmit(e) {
    e.preventDefault()

    let file = document.getElementById("upload-file").files[0];
    let fileType = file.type

    let formData = new FormData()
    formData.append("files", file)

    let url = null
    switch (fileType) {
        // .xlsx
        case 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
            url = '/xlsx'
            break;
        default:
            showResponse(uploadForm, 'Не известный тип файла!', true)
            console.log(`FileType is undefined: ${fileType}`);
            break;
    }

    if ( url ) {
        fetch(server_url + url, {
            method: 'POST',
            body: formData,
            headers: {
                "Content-Type": "multipart/form-data"
            }
        })
            .then(res => {
                switch (res.status) {
                    case 200:
                        showResponse(uploadForm, 'Файл успешно загружен!');
                        break;
                
                    default:
                        showResponse(uploadForm, 'Ошибка сервера', true)
                        console.log(`Server error (with response): ${res.status}`)
                        break;
                }
            })
            .catch(err => {
                showResponse(uploadForm, 'Ошибка сервера', true)
                console.log(`Server error: ${err}`)
            });
    }
}

function showResponse(form, message, isError = false) {
    let resElement = form.querySelector('p.form-response')
    if (!resElement) {
        resElement = document.createElement('p')
        form.appendChild(resElement)
    }

    if (isError) {
        resElement.className = 'form-response error'
    } else {
        resElement.className = 'form-response'
    }

    resElement.innerText = message
    
}