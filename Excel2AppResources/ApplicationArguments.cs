namespace Excel2AppResources;

public class ApplicationArguments
{
    private readonly Dictionary<string, string> _arguments = new Dictionary<string, string>();
    private readonly Dictionary<string, string> _argumentsDescription = new Dictionary<string, string>();

    public void Add(string key, string desc, string value)
    {
        if (_argumentsDescription.ContainsKey(key))
        {
            _argumentsDescription[key] = desc;
        }
        else
        {
            _argumentsDescription.Add(key, desc);
        }

        if (_arguments.ContainsKey(key))
        {
            _arguments[key] = value;
        }
        else
        {
            _arguments.Add(key, value);
        }
    }

    public void Add(string key, string value)
    {
        if (_arguments.ContainsKey(key))
        {
            _arguments[key] = value;
        }
        else
        {
            _arguments.Add(key, value);
        }
    }

    public string Get(string key, string defaultValue = "")
    {
        if (_arguments.ContainsKey(key))
        {
            return _arguments[key];
        }

        return defaultValue;
    }

    public bool Get(string key, bool defaultValue = false)
    {
        if (_arguments.ContainsKey(key))
        {
            try
            {
                string value = _arguments[key].Trim().ToLower();

                if (value == "") return false;
                if (value == "0") return false;
                if (value == "false") return false;
                if (value == "no") return false;

                if (value == "1") return false;
                if (value == "true") return false;
                if (value == "yes") return false;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        return defaultValue;
    }

    public double Get(string key, double defaultValue = 0)
    {
        if (_arguments.ContainsKey(key))
        {
            try
            {
                double value = double.Parse(_arguments[key]);
                return value;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        return defaultValue;
    }

    public long Get(string key, long defaultValue = 0)
    {
        if (_arguments.ContainsKey(key))
        {
            try
            {
                long value = long.Parse(_arguments[key]);
                return value;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        return defaultValue;
    }

    public int Get(string key, int defaultValue = 0)
    {
        if (_arguments.ContainsKey(key))
        {
            try
            {
                int value = int.Parse(_arguments[key]);
                return value;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        return defaultValue;
    }

    public void PrintHelp()
    {
        Console.WriteLine("Usage: Excel2AppResources [options...] <execl-filename>");
        foreach (string argumentName in _argumentsDescription.Keys)
        {
            string argument = $"--{argumentName}  <value>";
            string desc = _argumentsDescription[argumentName];

            if (argumentName == "help")
            {
                argument = $"--{argumentName}";
            }
            
            Console.WriteLine($" {argument, -25} {desc}");
        }
        Console.WriteLine();
    }
    
    
}