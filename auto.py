import os
import random
import time
from PIL import Image, ImageDraw, ImageFont
import pyexcel_ods3
import threading  # For running the stopwatch in parallel

# You don't need to adjust anything
def wrap_text(draw, text, font, max_width):
    """Helper function to wrap text into multiple lines based on image width."""
    lines = []
    words = text.split(' ')
    current_line = ''
    for word in words:
        test_line = current_line + ' ' + word if current_line else word
        bbox = draw.textbbox((0, 0), test_line, font=font)
        line_width = bbox[2] - bbox[0]
        if line_width <= max_width:
            current_line = test_line
        else:
            lines.append(current_line)
            current_line = word
    lines.append(current_line)
    return lines

# You don't need to adjust anything
def add_text_to_image(image_path, text, output_path, font_path, max_width_ratio=0.9, line_spacing_factor=1.5):
    """Function to add centered text to an image."""
    image = Image.open(image_path)
    draw = ImageDraw.Draw(image)
    font = ImageFont.truetype(font_path, 80)

    image_width, image_height = image.size
    max_text_width = int(image_width * max_width_ratio)

    lines = wrap_text(draw, text, font, max_text_width)

    total_text_height = sum(
        int((draw.textbbox((0, 0), line, font=font)[3] - draw.textbbox((0, 0), line, font=font)[
            1]) * line_spacing_factor)
        for line in lines
    )

    start_y = (image_height - total_text_height) // 2

    outline_thickness = 5
    outline_color = 'black'
    text_color = 'white'

    current_y = start_y
    for line in lines:
        text_bbox = draw.textbbox((0, 0), line, font=font)
        text_width = text_bbox[2] - text_bbox[0]
        position = ((image_width - text_width) // 2, current_y)

        # Draw outline
        for x_offset in range(-outline_thickness, outline_thickness + 1):
            for y_offset in range(-outline_thickness, outline_thickness + 1):
                if x_offset != 0 or y_offset != 0:
                    draw.text((position[0] + x_offset, position[1] + y_offset), line, font=font, fill=outline_color)

        # Draw the actual text
        draw.text(position, line, font=font, fill=text_color)

        line_height = text_bbox[3] - text_bbox[1]
        current_y += int(line_height * line_spacing_factor)

    image.save(output_path)

# You don't need to adjust anything
def real_time_stopwatch(start_time, running_flag):
    """Real-time stopwatch that runs in a separate thread and displays elapsed time."""
    while running_flag[0]:
        elapsed_time = time.time() - start_time
        print(f"\rElapsed Time: {elapsed_time:.2f} seconds", end="")
        time.sleep(1)  # Update every second
    # Print the final time when the process is done
    elapsed_time = time.time() - start_time
    print(f"\rDone in {elapsed_time:.2f} seconds")

# Adjust "repeat count" to how many you want to repeat the same text in the posts
def process_captions_and_create_images(captions_file, background_images_folder, output_folder, font_path,
                                       repeat_count=3):
    """Main function to process the captions and generate image packs multiple times."""

    # Start the stopwatch
    start_time = time.time()
    running_flag = [True]

    # Start the stopwatch thread
    stopwatch_thread = threading.Thread(target=real_time_stopwatch, args=(start_time, running_flag))
    stopwatch_thread.start()

    # Load the captions from the .ods file
    captions_data = pyexcel_ods3.get_data(captions_file)

    # Ensure output folder exists
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Get the list of background images
    background_images = [img for img in os.listdir(background_images_folder) if img.lower().endswith('.png')]
    background_images = [os.path.join(background_images_folder, img) for img in background_images]

    # Debugging information
    if not background_images:
        raise RuntimeError("No valid background images found in the specified folder.")

    print(f"Found {len(background_images)} background images.")

    for repeat_idx in range(repeat_count):  # Repeat process 5 times
        for i, row in enumerate(captions_data['Sheet1']):
            if not row:
                continue

            # Create a folder for the post pack (e.g., "post_1_repeat_1", "post_1_repeat_2", ...)
            post_folder = os.path.join(output_folder, f'post_{i + 1}_repeat_{repeat_idx + 1}')
            os.makedirs(post_folder, exist_ok=True)

            used_images = set()  # Track which images are already used in the pack

            for j, text in enumerate(row):
                # Choose a random background image that hasn't been used
                available_images = list(set(background_images) - used_images)

                if not available_images:
                    print(f"No more background images available for post {i + 1}, repeat {repeat_idx + 1}.")
                    break  # Exit the loop if no images are available

                image_path = random.choice(available_images)
                used_images.add(image_path)

                # Define the output image path (e.g., "1.png", "2.png", etc.)
                output_image_path = os.path.join(post_folder, f'{j + 1}.png')

                # Add text to the image and save it
                add_text_to_image(image_path, text, output_image_path, font_path)

    # Stop the real-time stopwatch thread
    running_flag[0] = False
    stopwatch_thread.join()


# Example usage with your paths and 5 repetitions
# Adjust the file directions
captions_file = 'captions.ods'
background_images_folder = 'backgrounds'
output_folder = 'output'
font_path = 'font.ttf'

# Call the function to repeat the process 5 times
# Adjust "repeat count" to how many you want to repeat the same text in the posts
process_captions_and_create_images(captions_file, background_images_folder, output_folder, font_path, repeat_count=3)