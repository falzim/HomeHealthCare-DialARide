#!/usr/local_rwth/bin/zsh

#how many nodes we want (1)
#SBATCH --nodes=1

### Start of Slurm SBATCH definitions
# give gams all 48  cores in a node
#SBATCH --ntasks=48

# give gams all the memory on the node
#SBATCH --mem=180G


# Name the job
#SBATCH --job-name=JOB_P1

#SBATCH --time=60:00:00
#SBATCH --account=rwth1638
##in case the account does not work, change to my account id


# Declare a file where the STDOUT/STDERR outputs will be written
# Define the output and error files
##file that is created if code ran successfully 
#SBATCH --output=/home/yh573838/outputs/out_unified_%J.out
##file that is created if code crashed
#SBATCH --error=/home/yh573838/outputs/error_unified_%J.err


###########################################

hostname
#free -g 
#memquota

module load GCCcore/.12.2.0
module load Gurobi/10.0.0-Python-3.11.2
#Gives system information output, use only when requested by someone
#echo $R_DELIMITER; export; echo $R_DELIMITER; echo


date
 python3 main.py

date
