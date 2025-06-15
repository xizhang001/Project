import pandas as pd
import pytesseract
from PIL import Image
import textract
from pdf2image import convert_from_path
from rapidfuzz import process, fuzz
import mimetypes
import os

# === Load and initialize ranking data ===
def init_ranking_data(excel_path, sheet_name):
    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=1)
    if 'Name of Institution' not in df.columns:
        raise ValueError("Ranking sheet must include 'Name of Institution' column.")
    institution_list = df['Name of Institution'].dropna().str.lower().tolist()
    return df, institution_list

# === Extract text from file (PDF: textract then OCR fallback) ===
def extract_text_from_file(file_path):
    file_type, _ = mimetypes.guess_type(file_path)
    print(f"\n🔍 Extracting text from: {file_path}")

    try:
        if file_path.lower().endswith(".pdf"):
            try:
                text = textract.process(file_path).decode('utf-8').strip()
                if len(text) > 50:
                    print("✅ Used textract for PDF")
                    print(text[:500])
                    return text
            except Exception:
                pass

            images = convert_from_path(file_path)
            text = ""
            for img in images:
                text += pytesseract.image_to_string(img)
            print("✅ Used OCR for PDF")
            print(text[:500])
            return text

        elif file_type and ("word" in file_type or "text" in file_type):
            text = textract.process(file_path).decode('utf-8')
            print("✅ Used textract for DOCX/TXT")
            print(text[:500])
            return text

        elif file_type and "image" in file_type:
            image = Image.open(file_path)
            text = pytesseract.image_to_string(image)
            print("✅ Used OCR for image")
            print(text[:500])
            return text

        print("⚠️ Unsupported or unknown file type")
        return ""

    except Exception as e:
        print(f"❌ Error processing file {file_path}: {e}")
        return ""

# === Fuzzy match with option to return score ===
def find_institution(text, institution_list, return_score=False):
    # Preprocess text - handle multi-line names and OCR artifacts
    cleaned_text = ' '.join(text.lower().split())
    normalized_text = cleaned_text.replace('-', ' ').replace('.', '')
    
    # Print for debugging
    print(f"\n=== DEBUG: Normalized Text ===")
    print(normalized_text[:500])
    
    # Strategy 1: Check for exact multi-line matches
    for inst in institution_list:
        # Create variations: with and without "Dr." prefix
        variations = [
            inst.lower(),
            inst.lower().replace("dr. ", "").strip(),
            inst.lower().replace("dr ", "").strip()
        ]
        
        for variant in variations:
            # Check if all words of the variant appear in the text
            variant_words = variant.split()
            if all(word in normalized_text for word in variant_words):
                print(f"Exact match found: {inst} (Variant: {variant})")
                return (inst, 100) if return_score else inst
    
    # Strategy 2: Fuzzy matching with word components
    best_match = None
    best_score = 0
    
    for inst in institution_list:
        # Create search pattern from institution name components
        search_terms = [
            term for term in inst.lower().split() 
            if len(term) > 3  # Ignore short words
        ]
        
        # Calculate match score based on presence of key terms
        match_score = sum(1 for term in search_terms if term in normalized_text)
        completeness = match_score / len(search_terms) if search_terms else 0
        
        # Calculate additional score for ordered appearance
        ordered_score = 0
        if search_terms:
            try:
                positions = [normalized_text.index(term) for term in search_terms]
                if positions == sorted(positions):
                    ordered_score = 0.3  # Bonus for correct order
            except ValueError:
                pass
        
        total_score = (completeness + ordered_score) * 100
        
        # Track best match
        if total_score > best_score:
            best_score = total_score
            best_match = inst
            print(f"New best match: {inst} (Score: {total_score:.1f})")
    
    # Apply threshold
    if best_score >= 75:
        print(f"Best match accepted: {best_match} ({best_score:.1f}%)")
        return (best_match, best_score) if return_score else best_match
    
    print("No strong match found")
    return (None, best_score) if return_score else None

# === Lookup ranking info ===
def lookup_institution_ranking(name, ranking_df, extracted_text=None):
    name = name.lower()
    matches = ranking_df[ranking_df['Name of Institution'].str.lower() == name]

    if matches.empty:
        return None

    if len(matches) == 1:
        row = matches.iloc[0]
    else:
        filtered = matches.copy()
        if extracted_text:
            extracted_text = extracted_text.lower()
            if 'City' in ranking_df.columns:
                filtered = filtered[filtered['City'].astype(str).str.lower().apply(lambda c: c in extracted_text)]
            if 'State' in ranking_df.columns and len(filtered) > 1:
                filtered = filtered[filtered['State'].astype(str).str.lower().apply(lambda s: s in extracted_text)]
        row = filtered.iloc[0] if not filtered.empty else matches.iloc[0]

    base_info = {
        "Name of Institution": row["Name of Institution"],
        "City": row.get("CITY") or row.get("City"),
        "State": row.get("STATE") or row.get("State"),
    }

    tier_1_cols = [
        "Top 100 Overall", "Top 100 University", "Top 100 College",
        "Top 100 Engineering", "QS Global"
    ]
    tier_2_cols = [
        "101-200 Overall", "101-200 University", "101-200 College"
    ]

    tier_1 = {col: row[col] for col in tier_1_cols if pd.notna(row.get(col)) and str(row[col]).strip()}
    tier_2 = {col: row[col] for col in tier_2_cols if pd.notna(row.get(col)) and str(row[col]).strip()}

    result = base_info
    if tier_1:
        result["Tier 1"] = tier_1
    if tier_2:
        result["Tier 2"] = tier_2

    return result

# === Main logic ===
def process_student_files(transcript_path=None, cv_path=None, reference_paths=None, ranking_df=None, institution_list=None):
    source = None
    matched_name = None
    match_score = 0
    raw_text = ""

    # Step 1: Transcript
    if transcript_path:
        text = extract_text_from_file(transcript_path)
        raw_text = text  # Always store text
        name, score = find_institution(text, institution_list, return_score=True)
        if name:
            matched_name = name
            match_score = score
            source = "Transcript"

    # Step 2: CV
    if not matched_name and cv_path:
        text = extract_text_from_file(cv_path)
        raw_text = text  # Always store text
        name, score = find_institution(text, institution_list, return_score=True)
        if name:
            matched_name = name
            match_score = score
            source = "CV"

    # Step 3: References
    if not matched_name and reference_paths:
        for path in reference_paths:
            text = extract_text_from_file(path)
            raw_text = text  # Always store text, even on final iteration
            name, score = find_institution(text, institution_list, return_score=True)
            if name:
                matched_name = name
                match_score = score
                source = f"Reference Letter: {os.path.basename(path)}"
                break

    # Step 4: No match — return at least extracted info
    if not matched_name:
        return {
            "_raw_text": raw_text,
            "_source": source if source else "No match",
            "_match_score": match_score
        }

    # Step 5: Lookup rankings
    result = lookup_institution_ranking(matched_name, ranking_df, extracted_text=raw_text)
    if result:
        result["_source"] = source
        result["_raw_text"] = raw_text
        result["_match_score"] = match_score
    return result
