using System.Net;
using System.Text;
using System.Windows.Forms;

namespace Phone_EXP_Builder
{
    public partial class Form1 : Form
    {
        private Dictionary<string, string> nameToNumberDictionary;
        //private int maxPairsPerTab;

        public Form1()
        {
            InitializeComponent();
            GenerateOutputButton.Click += GenerateOutputButton_Click;
            nameToNumberDictionary = new Dictionary<string, string>();
            string csvUrl = "https://raw.githubusercontent.com/Banana-Goats/Ableworld/main/Contacts.csv";
            PopulateDictionaryFromCsvAsync(csvUrl);

        }

        private void GenerateOutputButton_Click(object sender, EventArgs e)
        {
            // Initialize the output string with pre-chosen text
            StringBuilder outputBuilder = new StringBuilder();

            // Get the maximum number of tabs in the TabControl
            int maxTabs = tabControl1.TabCount;

            // Iterate through tabs
            for (int tabIdx = 0; tabIdx < maxTabs; tabIdx++)
            {
                TabPage tabPage = tabControl1.TabPages[tabIdx];

                // Get the maximum number of pairs for this tab (e.g., 60 pairs per tab)
                int maxPairsPerTab = 60;

                // Iterate through the pairs on this tab
                for (int i = 1; i <= maxPairsPerTab; i++)
                {
                    // Attempt to get the values from the corresponding text boxes
                    string name = null;
                    string number = null;

                    try
                    {
                        name = tabPage.Controls[$"name{i}"]?.Text;
                        number = tabPage.Controls[$"number{i}"]?.Text;
                    }
                    catch (Exception ex)
                    {
                        // Handle any exceptions (e.g., control not found)
                        // You can log the exception if needed
                        Console.WriteLine($"An exception occurred: {ex.Message}");
                    }

                    // Check if both name and number are filled for this pair
                    if (!string.IsNullOrWhiteSpace(name) && !string.IsNullOrWhiteSpace(number))
                    {
                        // Append the output for the current pair with the correct key number
                        outputBuilder.AppendLine($"expansion_module.1.key.{i}.extension = *04");
                        outputBuilder.AppendLine($"expansion_module.1.key.{i}.label = {name}");
                        outputBuilder.AppendLine($"expansion_module.1.key.{i}.line = 1");
                        outputBuilder.AppendLine($"expansion_module.1.key.{i}.type = 16");
                        outputBuilder.AppendLine($"expansion_module.1.key.{i}.value = {number}");
                        outputBuilder.AppendLine();
                    }
                }
            }


            // Set the generated output in the multiline text box
            outputTextBox.Text = outputBuilder.ToString();
        }

        private void Copybutton_Click(object sender, EventArgs e)
        {
            // Check if the TextBox has text
            if (!string.IsNullOrWhiteSpace(outputTextBox.Text))
            {
                // Copy the TextBox text to the clipboard
                Clipboard.SetText(outputTextBox.Text);
                MessageBox.Show("Text copied to clipboard!");
            }
            else
            {
                MessageBox.Show("TextBox is empty. Nothing to copy.");
            }
        }

        private void LookupNumbersButton_Click(object sender, EventArgs e)
        {
            int maxPairsPerTab = 60;

            foreach (Control control in tabControl1.Controls)
            {
                if (control is TabPage tabPage)
                {
                    for (int i = 1; i <= maxPairsPerTab; i++)
                    {
                        string nameTextBoxName = $"name{i}";
                        string numberTextBoxName = $"number{i}";

                        // Attempt to get the values from the corresponding text boxes
                        string name = tabPage.Controls[nameTextBoxName]?.Text;

                        if (!string.IsNullOrWhiteSpace(name))
                        {
                            if (nameToNumberDictionary.TryGetValue(name, out string number))
                            {
                                tabPage.Controls[numberTextBoxName].Text = number;
                            }
                            else
                            {
                                MessageBox.Show($"No number found for name: {name}");
                            }
                        }
                    }
                }
            }
        }
        private void PopulateDictionary()
        {
            // Clear the dictionary to start fresh
            nameToNumberDictionary.Clear();

            // Path to your CSV file
            string csvFilePath = "C:\\\\Users\\erenner\\Downloads\\Contacts.csv";

            try
            {
                // Read all lines from the CSV file
                string[] lines = File.ReadAllLines(csvFilePath);

                foreach (string line in lines)
                {
                    string[] parts = line.Split(',');

                    if (parts.Length == 2)
                    {
                        string name = parts[0].Trim();
                        string number = parts[1].Trim();
                        nameToNumberDictionary[name] = number;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while reading the CSV file: {ex.Message}");
            }
        }

        private async Task PopulateDictionaryFromCsvAsync(string csvUrl)
        {
            // Clear the dictionary to start fresh
            nameToNumberDictionary.Clear();

            try
            {
                using (HttpClient client = new HttpClient())
                {
                    // Download the CSV content
                    string csvContent = await client.GetStringAsync(csvUrl);

                    // Parse CSV content
                    string[] lines = csvContent.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);

                    foreach (string line in lines)
                    {
                        string[] parts = line.Split(',');

                        if (parts.Length == 2)
                        {
                            string name = parts[0].Trim();
                            string number = parts[1].Trim();
                            nameToNumberDictionary[name] = number;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while reading the CSV from {csvUrl}: {ex.Message}");
            }
        }
    }
}
