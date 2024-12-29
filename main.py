# // script
# // dependencies = ['nanodjango']
# //

from nanodjango import Django
from django.shortcuts import render, redirect
from django.http import HttpResponse
from CreateBill import create_bill
import io

app = Django()

@app.route('/')
def index(request):
	return render(request, 'index.html')


@app.route('/docx')
def docx(request):
    if request.method == 'POST':
        bill_no=request.POST.get('bill_no')
        showroom=request.POST.get('showroom')
        invoice = create_bill(
            billNo=bill_no,
            brand = request.POST.get('brand'),
            showroom = showroom,
            address = request.POST.get('address'),
            mobile = request.POST.get('mobile'),
            board = request.POST.get('board')
        )

        buffer = io.BytesIO()
        invoice.save(buffer)
        buffer.seek(0)
		
        response = HttpResponse(
            buffer, 
            content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
        
        response['Content-Disposition'] = f'attachment; filename="{bill_no} {showroom}.docx"'
        return response

    else:
        return render('new.html')



if __name__ == "__main__":
    app.run(host='0.0.0.0:8000')
