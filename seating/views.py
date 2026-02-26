from django.shortcuts import render, redirect
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.http import HttpResponse
from django.core.mail import send_mail
from django.conf import settings
from django.core.files.storage import FileSystemStorage
import pandas as pd
import random
import csv

from .models import Block, Room, ExamSlot


# =========================
# BLOCKS VIEW
# =========================
@login_required(login_url='login')
def blocks_view(request):
    if request.method == 'POST':

        if 'add' in request.POST:
            name = request.POST.get('block_name')
            if name:
                Block.objects.get_or_create(name=name)

        if 'delete' in request.POST:
            block_id = request.POST.get('delete')
            Block.objects.filter(id=block_id).delete()

        if 'next' in request.POST:
            selected = request.POST.getlist('blocks')
            request.session['selected_blocks'] = selected
            return redirect('rooms')

    blocks = Block.objects.all()
    return render(request, 'seating/blocks.html', {'blocks': blocks})


# =========================
# ROOMS VIEW
# =========================
@login_required(login_url='login')
def rooms_view(request):
    selected_blocks = request.session.get('selected_blocks', [])

    if request.method == 'POST':

        if 'add' in request.POST:
            block_id = request.POST.get('block')
            room_number = request.POST.get('room_number')
            rows = request.POST.get('rows')
            columns = request.POST.get('columns')

            if block_id and room_number and rows and columns:
                Room.objects.create(
                    block_id=block_id,
                    room_number=room_number,
                    rows=int(rows),
                    columns=int(columns),
                    capacity=int(rows) * int(columns)
                )
                return redirect('rooms')

        if 'delete' in request.POST:
            room_id = request.POST.get('delete')
            Room.objects.filter(id=room_id).delete()
            return redirect('rooms')

        if 'next' in request.POST:
            selected_rooms = request.POST.getlist('rooms')
            request.session['selected_rooms'] = selected_rooms
            return redirect('exam_slot')

    rooms = Room.objects.filter(block__id__in=selected_blocks)
    blocks = Block.objects.filter(id__in=selected_blocks)

    return render(request, 'seating/rooms.html', {
        'rooms': rooms,
        'blocks': blocks
    })


# =========================
# EXAM SLOT VIEW
# =========================
@login_required(login_url='login')
def exam_slot_view(request):
    if request.method == 'POST':

        if 'add' in request.POST:
            name = request.POST.get('name')
            date = request.POST.get('date')
            time_range = request.POST.get('time_range')

            if name and date and time_range:
                ExamSlot.objects.create(
                    name=name,
                    date=date,
                    time_range=time_range
                )

        if 'delete' in request.POST:
            slot_id = request.POST.get('delete')
            ExamSlot.objects.filter(id=slot_id).delete()

        if 'next' in request.POST:
            slot_id = request.POST.get('slot')
            request.session['exam_slot'] = slot_id
            return redirect('upload')

    slots = ExamSlot.objects.all()
    return render(request, 'seating/exam_slot.html', {'slots': slots})


# =========================
# UPLOAD EXCEL VIEW
# =========================
@login_required(login_url='login')
def upload_view(request):
    if request.method == 'POST' and request.FILES.get('excel'):
        excel_file = request.FILES['excel']

        fs = FileSystemStorage()
        filename = fs.save(excel_file.name, excel_file)
        file_path = fs.path(filename)

        df = pd.read_excel(file_path)

        # Store data in session
        request.session['students_data'] = df.to_dict(orient='records')

        return redirect('seating')

    return render(request, 'seating/upload.html')


# =========================
# SEATING VIEW (FIXED)
# =========================
@login_required(login_url='login')
def seating_view(request):

    students = request.session.get('students_data', [])
    room_ids = request.session.get('selected_rooms', [])

    rooms = Room.objects.filter(id__in=room_ids)

    if not students or not rooms:
        return render(request, 'seating/seating.html', {'seat_map': {}})

    # Group by Branch
    branch_map = {}
    for s in students:
        branch = s.get('Branch')
        branch_map.setdefault(branch, []).append(s)

    # Sort by Roll No (IMPORTANT FIX)
    for b in branch_map:
        branch_map[b] = sorted(
            branch_map[b],
            key=lambda x: str(x.get('Roll No', ''))
        )

    # Interleave students branch-wise
    ordered_students = []
    while any(branch_map.values()):
        for b in list(branch_map.keys()):
            if branch_map[b]:
                ordered_students.append(branch_map[b].pop(0))

    seating = []
    seat_map = {}
    idx = 0

    for room in rooms:
        room_seats = []

        for row in range(1, room.rows + 1):
            for col in range(1, room.columns + 1):

                if idx < len(ordered_students):
                    s = ordered_students[idx]

                    seat = {
                        'room': f"{room.block.name}{room.room_number}",
                        'row': row,
                        'col': col,
                        'roll': s.get('Roll No'),  # FIXED HERE
                        'name': s.get('Name'),
                        'email': s.get('Email'),
                        'branch': s.get('Branch'),
                        'subject': s.get('Subject')
                    }

                    room_seats.append(seat)
                    seating.append(seat)
                    idx += 1

        seat_map[f"{room.block.name}{room.room_number}"] = room_seats

    request.session['final_seating'] = seating

    return render(request, 'seating/seating.html', {
        'seat_map': seat_map
    })


# =========================
# DOWNLOAD SEATING
# =========================
@login_required(login_url='login')
def download_seating(request):
    seating = request.session.get('final_seating', [])

    if not seating:
        return HttpResponse("No seating data found.", status=400)

    df = pd.DataFrame(seating)

    df = df.rename(columns={
        'room': 'Room',
        'row': 'Row',
        'col': 'Column',
        'roll': 'Roll No',
        'name': 'Name',
        'email': 'Email',
        'branch': 'Branch',
        'subject': 'Subject'
    })

    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = 'attachment; filename=seating.xlsx'

    with pd.ExcelWriter(response, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Seating')

    return response


# =========================
# SEND EMAILS
# =========================
def send_seat_email(student, exam_slot):

    subject = "Your Exam Seating Details"

    message = f"""
Dear {student['name']},

Your exam seating has been allocated successfully.

Subject: {student['subject']}
Room: {student['room']}
Row: {student['row']}
Column: {student['col']}

Exam Date: {exam_slot.date}
Exam Time: {exam_slot.time_range}

Please report 30 minutes early.

All the best!
Exam Cell
"""

    send_mail(
        subject,
        message,
        settings.DEFAULT_FROM_EMAIL,
        [student['email']],
        fail_silently=False,
    )


def send_mails_view(request):

    seating = request.session.get('final_seating')
    slot_id = request.session.get('exam_slot')

    if not seating or not slot_id:
        messages.error(request, "No seating data found. Please generate seating first.")
        return redirect('seating')

    exam_slot = ExamSlot.objects.get(id=slot_id)

    for s in seating:
        send_seat_email(s, exam_slot)

    messages.success(request, "Emails sent successfully to all students!")
    return redirect('seating')


# =========================
# HOME
# =========================
@login_required(login_url='login')
def home(request):
    return render(request, 'seating/home.html')