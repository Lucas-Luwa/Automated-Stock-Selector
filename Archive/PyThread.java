import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.InputStream;

class PyThread extends Thread {
    private String funcName, methodName, pythonFileName;
    
    public PyThread(String funcName, String pythonFileName, String methodName) {
        super(funcName);
        this.pythonFileName = pythonFileName;
        this.methodName = methodName;
    }
    
    @Override
    public void run() {
        try {
            ProcessBuilder procBuilder = new ProcessBuilder("python", pythonFileName, methodName);
            Process myProc = procBuilder.start();
            BufferedReader reader = new BufferedReader(new InputStreamReader(myProc.getInputStream()));
            String tempHolder;
            StringBuilder output = new StringBuilder();
            while ((tempHolder = reader.readLine()) != null) {
                output.append(tempHolder).append("\n");
            }

            int exitCode = myProc.waitFor();
            System.out.println("Thread: " + getName());
            System.out.println("Output:\n" + output);
            System.out.println("Exit Code: " + exitCode);    
        } catch (Exception e){
            e.printStackTrace();
        }
    }
}