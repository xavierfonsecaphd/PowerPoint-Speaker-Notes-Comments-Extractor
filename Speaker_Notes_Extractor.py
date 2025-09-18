import win32com.client
import sys
import os
import json
from datetime import datetime

def extract_speaker_notes(ppt_path):
    """
    Extract speaker notes from PowerPoint presentation
    
    Args:
        ppt_path (str): Path to PowerPoint file
        
    Returns:
        list: List of notes dictionaries
    """
    notes = []
    ppt_app = None
    presentation = None
    
    try:
        # Start PowerPoint
        print("Starting PowerPoint application...")
        ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        print("✓ PowerPoint started")
        
        # Open presentation
        full_path = os.path.abspath(ppt_path)
        presentation = ppt_app.Presentations.Open(full_path)
        print(f"✓ Opened: {presentation.Name}")
        print(f"✓ Total slides: {presentation.Slides.Count}")
        
        # Extract notes from each slide
        for slide_num in range(1, presentation.Slides.Count + 1):
            slide = presentation.Slides(slide_num)
            slide_notes = []
            
            try:
                # Access the notes page
                if hasattr(slide, 'NotesPage'):
                    notes_page = slide.NotesPage
                    
                    # Look through all shapes on the notes page
                    for shape_idx in range(1, notes_page.Shapes.Count + 1):
                        shape = notes_page.Shapes(shape_idx)
                        
                        # Check if shape has text
                        if hasattr(shape, 'TextFrame') and shape.TextFrame.HasText:
                            try:
                                notes_text = shape.TextFrame.TextRange.Text.strip()
                                
                                # Filter out empty or very short text
                                if notes_text and len(notes_text) > 3:
                                    # Skip if it's just slide title repetition (common in notes)
                                    if not notes_text.isdigit() and notes_text not in ['Slide', 'Notes']:
                                        slide_notes.append({
                                            'shape_index': shape_idx,
                                            'text': notes_text,
                                            'shape_type': shape.Type if hasattr(shape, 'Type') else 'Unknown'
                                        })
                                        
                            except Exception as e:
                                print(f"    Error reading shape {shape_idx}: {e}")
                                continue
                
                # If we found notes on this slide
                if slide_notes:
                    print(f"✓ Slide {slide_num}: Found {len(slide_notes)} notes")
                    for note in slide_notes:
                        preview = note['text'][:100] + "..." if len(note['text']) > 100 else note['text']
                        print(f"    {preview}")
                    
                    notes.append({
                        'slide_number': slide_num,
                        'notes_count': len(slide_notes),
                        'notes': slide_notes
                    })
                else:
                    print(f"  Slide {slide_num}: No notes")
                    
            except Exception as e:
                print(f"  Error processing slide {slide_num}: {e}")
                continue
        
        print(f"\n✓ Total slides with notes: {len(notes)}")
        
    except Exception as e:
        print(f"Error: {str(e)}")
        
    finally:
        # Clean up
        try:
            if presentation:
                presentation.Close()
            if ppt_app:
                ppt_app.Quit()
            print("✓ PowerPoint closed")
        except:
            pass
    
    return notes

def save_notes(notes, output_base_name):
    """Save speaker notes to files"""
    
    if not notes:
        print("No speaker notes to save.")
        return
    
    # Count total notes
    total_notes = sum(slide_data['notes_count'] for slide_data in notes)
    
    # Save as JSON
    json_file = f"{output_base_name}_speaker_notes.json"
    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(notes, f, indent=2, ensure_ascii=False, default=str)
    print(f"✓ Speaker notes saved to: {json_file}")
    
    # Save as text file
    txt_file = f"{output_base_name}_speaker_notes.txt"
    with open(txt_file, 'w', encoding='utf-8') as f:
        f.write("PowerPoint Speaker Notes Extract\n")
        f.write("=" * 50 + "\n\n")
        f.write(f"Total slides with notes: {len(notes)}\n")
        f.write(f"Total note items: {total_notes}\n")
        f.write(f"Extraction date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        
        for slide_data in notes:
            slide_num = slide_data['slide_number']
            f.write(f"SLIDE {slide_num}\n")
            f.write("=" * 20 + "\n")
            
            for i, note in enumerate(slide_data['notes'], 1):
                f.write(f"\nNote #{i} (Shape {note['shape_index']}):\n")
                f.write(f"Type: {note['shape_type']}\n")
                f.write(f"Text:\n{note['text']}\n")
                f.write("-" * 30 + "\n")
            
            f.write("\n\n")
    
    print(f"✓ Speaker notes saved to: {txt_file}")

def main():
    """Main function"""
    if len(sys.argv) != 2:
        print("Usage: python script.py <powerpoint_file.pptx>")
        print("Example: python script.py presentation.pptx")
        sys.exit(1)
    
    ppt_file = sys.argv[1]
    
    # Check if file exists
    if not os.path.exists(ppt_file):
        print(f"Error: File '{ppt_file}' not found.")
        sys.exit(1)
    
    print(f"Extracting speaker notes from: {ppt_file}")
    print("=" * 60)
    
    # Extract notes
    notes = extract_speaker_notes(ppt_file)
    
    if notes:
        # Display summary
        total_notes = sum(slide_data['notes_count'] for slide_data in notes)
        
        print(f"\n📋 EXTRACTION SUMMARY")
        print("=" * 30)
        print(f"Slides with notes: {len(notes)}")
        print(f"Total note items: {total_notes}")
        
        # Show which slides have notes
        print(f"\nSlides containing notes:")
        for slide_data in notes:
            print(f"  Slide {slide_data['slide_number']}: {slide_data['notes_count']} note(s)")
        
        # Show preview of first few notes
        print(f"\n📄 NOTES PREVIEW")
        print("=" * 30)
        preview_count = 0
        for slide_data in notes:
            if preview_count >= 5:  # Show first 5 notes
                break
            for note in slide_data['notes']:
                if preview_count >= 5:
                    break
                preview_count += 1
                preview_text = note['text'][:150] + "..." if len(note['text']) > 150 else note['text']
                print(f"{preview_count}. Slide {slide_data['slide_number']}:")
                print(f"   {preview_text}")
                print()
        
        if total_notes > 5:
            print(f"... and {total_notes - 5} more notes")
        
        # Save to files
        output_base = os.path.splitext(ppt_file)[0]
        save_notes(notes, output_base)
        
    else:
        print("\n❌ No speaker notes found in the presentation.")

# Simple function for direct use
def get_speaker_notes(ppt_path):
    """
    Simple function to get speaker notes - for use in other scripts
    
    Args:
        ppt_path (str): Path to PowerPoint file
        
    Returns:
        list: List of notes dictionaries
    """
    return extract_speaker_notes(ppt_path)

if __name__ == "__main__":
    main()