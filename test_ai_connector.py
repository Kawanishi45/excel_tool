"""
AI連携モジュールのテストスクリプト
"""
import ai_connector
import os


def main():
    JSON_PATH = "output/instructions.json"
    IMAGE_PATH = "output/anchor_image.png"

    print("Testing ai_connector...")
    print("=" * 60)

    # Check if files exist
    if not os.path.exists(JSON_PATH):
        print(f"✗ Error: {JSON_PATH} not found")
        print("Please run test_asset_gen.py first to generate test files")
        return

    if not os.path.exists(IMAGE_PATH):
        print(f"✗ Error: {IMAGE_PATH} not found")
        print("Please run test_asset_gen.py first to generate test files")
        return

    print(f"✓ JSON file found: {JSON_PATH}")
    print(f"✓ Image file found: {IMAGE_PATH}")

    # Test prompt building
    print("\n[Step 1] Building prompt...")
    try:
        prompt_text, image_object = ai_connector.build_prompt(JSON_PATH, IMAGE_PATH)

        print(f"✓ Prompt generated ({len(prompt_text)} characters)")
        print(f"✓ Image loaded ({image_object.size})")

        print("\n--- Prompt Preview (first 500 chars) ---")
        print(prompt_text[:500])
        print("...")
        print("\n" + "=" * 60)

    except Exception as e:
        print(f"✗ Error building prompt: {e}")
        import traceback
        traceback.print_exc()
        return

    # Test API call (only if API key is configured)
    print("\n[Step 2] Testing API call...")
    api_key = os.environ.get('GOOGLE_API_KEY')

    if not api_key or api_key == 'your_api_key_here':
        print("⚠ GOOGLE_API_KEY not configured")
        print("To test API call:")
        print("  1. Get API key from: https://makersuite.google.com/app/apikey")
        print("  2. Create .env file: cp .env.example .env")
        print("  3. Edit .env and set your API key")
        print("\n✓ Prompt building test passed")
    else:
        print(f"✓ API key found (length: {len(api_key)})")
        print("Calling Gemini API...")

        try:
            mermaid_code = ai_connector.generate_mermaid_code(JSON_PATH, IMAGE_PATH)

            print("\n✓ Mermaid code generated!")
            print("=" * 60)
            print(mermaid_code)
            print("=" * 60)

            # Save to file
            output_file = "output/generated_mermaid.md"
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write("```mermaid\n")
                f.write(mermaid_code)
                f.write("\n```\n")

            print(f"\n✓ Saved to: {output_file}")

        except Exception as e:
            print(f"\n✗ API call failed: {e}")
            import traceback
            traceback.print_exc()

    print("\n" + "=" * 60)
    print("✓ Test complete!")


if __name__ == "__main__":
    main()
