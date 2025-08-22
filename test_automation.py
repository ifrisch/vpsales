"""
Test Van Paper Automation
Quick test to verify the automation works with your current email
"""

import subprocess
import sys
from pathlib import Path

def test_automation():
    """Test the scheduled automation script"""
    
    print("🧪 Testing Van Paper Automation")
    print("=" * 40)
    
    current_dir = Path(__file__).parent
    automation_script = current_dir / "scheduled_automation.py"
    
    if not automation_script.exists():
        print("❌ scheduled_automation.py not found!")
        return False
    
    print("🚀 Running automation test...")
    print("-" * 40)
    
    try:
        # Run the automation script
        result = subprocess.run([sys.executable, str(automation_script)], 
                              cwd=current_dir, 
                              capture_output=False,  # Show output in real-time
                              text=True)
        
        print("-" * 40)
        
        if result.returncode == 0:
            print("✅ Automation test completed successfully!")
            return True
        else:
            print(f"⚠️ Automation completed with return code: {result.returncode}")
            return False
            
    except Exception as e:
        print(f"❌ Error running automation: {e}")
        return False

if __name__ == "__main__":
    success = test_automation()
    
    print("\n" + "=" * 40)
    if success:
        print("🎉 Test passed! Automation is ready for scheduling.")
        print("\n📝 Next steps:")
        print("1. Run setup_scheduler.ps1 as Administrator to create scheduled tasks")
        print("2. Tasks will run at 7:32 AM and 2:02 PM CST automatically")
        print("3. Monitor the live app: https://vpsales.streamlit.app/")
    else:
        print("⚠️ Test had issues. Check the output above.")
    
    input("\nPress Enter to continue...")
