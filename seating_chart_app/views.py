import cv2
from django.http import StreamingHttpResponse
from django.shortcuts import render
from .makeup_processor import MakeupApplication  # Import the MakeupApplication class

# Initialize the makeup application instance
makeup_app = MakeupApplication()

# This view renders the index.html page
def index(request):
    # Renders the template for the homepage
    return render(request, 'seating_chart_app/upload.html')

# This view streams the video feed processed by the makeup application
def video_feed(request):
    # Returns a streaming response that will display the processed video feed
    return StreamingHttpResponse(gen(makeup_app),
                                 content_type='multipart/x-mixed-replace; boundary=frame')

# Generator function to read frames, process them, and serve as a video stream
def gen(makeup_app):
    cap = cv2.VideoCapture(0)  # Capture video from the webcam
    while cap.isOpened():
        success, frame = cap.read()  # Read a frame from the webcam
        if not success:
            break

        # Process the frame using the makeup application (eyeshadow, eyeliner, lipstick, etc.)
        frame = makeup_app.process_frame(frame)

        # Encode the frame in JPEG format
        ret, buffer = cv2.imencode('.jpg', frame)
        frame = buffer.tobytes()  # Convert the frame to bytes

        # Yield the frame to the HTTP response to be displayed on the web page
        yield (b'--frame\r\n'
               b'Content-Type: image/jpeg\r\n\r\n' + frame + b'\r\n')

    cap.release()  # Release the webcam
