from environs import Env

env = Env()
env.read_env()

ad_password = env.str('AD_PASSWORD')
ad_user = env.str('AD_USER')
ad_search_tree = env.str('AD_SEARCH_TREE')
server = env.str('SERVER')
