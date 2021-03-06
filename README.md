# Fair-Price-Per-Share-Calculator
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


<b>What it does:</b>
1. Calculate Value Drivers
2. Estimate Weighted Average Cost of Capital 
3. Project future potential financial performance 
    -Estimates share price based on long term growth potential (default = 1.25%) of yearly free cash flows

<b>To Use:</b>
1. download package
2. type in desired ticker symbol (line 13 on main.py)
3. run main.py using a Python 3.5

<b>How it works:</b>
1. This script will download 5 years financial data
2. Create a directory with the ticker symbol name
3. Load financial statements to CSV format
4. Calculate value drivers
5. Project valuation on excel for desired amount of years. (line 35 last parameter denotes years_to_project)

<b>TO USE:</b> 
1. Download and extract files.
2. Install required packages:
    pip install -r [Folder Path]/Fair-Price-Per-Share-Calculator/requirments.txt

3. <b>Run main file:</b>
    python <i>[Folder Path]</i>/Fair-Price-Per-Share-Calculator/main.py

4. Enter the ticker desired ticker symbol.

5. Data will be extracted to the script folder by ticker symbol.

<b>Warnings about the data:</b> 
1. Depreciation is not calculated... because most companies do not share "Gross, property plant and equipment" on their balance sheets. It is defaulted at 40%, much higher than most companies actually depreciate their equipment. You can calculate this by dividing total "Depreciation" from cash flow sheet by "Gross, property plant and equipment" usually provided in the actual annual 10k.


<b>Pending  Working on:</b>
1. Refactor the excel data injection methods. Pretty messy and convoluted.
2. Data calls directly to the dict structs need to be removed and encapsulated for error handling.
3. Comments need to be added to each class, functions, and variables within
4. Being a better person.
