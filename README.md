# Fair-Price-Per-Share-Calculator
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


What it does:
1. Calculate Value Drivers
2. Estimate WACC (Done, Shares outstanding calculation is rough, based on Market Cap/Share Price.)
3. Project future potential financial performance (Estimates share price based on long term growth potential default = 1.25%)

To Use:
1. download package
2. type in desired ticker symbol (line 13 on main.py)
3. run main.py using a Python 3.5

How it works:
1. This script will download 5 years financial data
2. Create a directory with the ticker symbol name
3. Load financial statements to CSV format
4. Calculate value drivers
5. Project valuation on excel for desired amount of years. (line 35 last parameter denotes years_to_project)
